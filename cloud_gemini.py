import os
import re
import time
import json
import uuid
import dotenv
import logging
import shutil
import pandas as pd
import openpyxl
import streamlit as st
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError as GoogleHttpError # Specific import
import google.generativeai as genai
import yt_dlp
import boto3
from botocore.exceptions import ClientError, NoCredentialsError
from typing import Dict, List, Optional, Tuple, Any

# --- Configuration ---

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(threadName)s %(message)s", # Added threadName
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Load environment variables
dotenv.load_dotenv()

# Constants
MAX_RESULTS_PER_PAGE = 50
SLEEP_TIME_BETWEEN_ANALYSES = 15  # Increased slightly for potential API limits
RETRY_ATTEMPTS = 3
BACKOFF_FACTOR = 2
MAX_BG_WORKERS = 1 # Limit concurrency for t2.micro
TEMP_PROCESSING_DIR = Path("./temp_processing") # Base temp directory

# API Keys
YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
GENAI_API_KEY = os.getenv("GENAI_API_KEY")

# S3 Configuration
AWS_REGION = os.getenv("AWS_REGION", "us-east-1") # Specify region
S3_BUCKET_NAME = os.getenv("AWS_BUCKET_NAME") # Main bucket for videos/tasks
S3_VIDEOS_FOLDER = os.getenv("S3_VIDEOS_FOLDER", "youtube_videos") # Folder for raw videos
S3_TASKS_FOLDER = os.getenv("S3_TASKS_FOLDER", "background_tasks") # Folder for task status
S3_OUTPUT_BUCKET_NAME = os.getenv("S3_OUTPUT_BUCKET_NAME", "youtube-summarizer-output-files") # Bucket for final results

# --- Global Clients and Executor ---
try:
    s3_client = boto3.client("s3", region_name=AWS_REGION)
    # Check S3 access at startup
    s3_client.head_bucket(Bucket=S3_BUCKET_NAME)
    s3_client.head_bucket(Bucket=S3_OUTPUT_BUCKET_NAME)
    logger.info(f"Successfully connected to S3 buckets: {S3_BUCKET_NAME}, {S3_OUTPUT_BUCKET_NAME}")
except (ClientError, NoCredentialsError) as e:
    logger.error(f"S3 connection/access error: {e}. Please check credentials and bucket names/permissions.")
    st.error(f"S3 connection/access error: {e}. Cannot proceed.")
    st.stop() # Stop the streamlit app if S3 isn't configured
except Exception as e:
    logger.error(f"Unexpected error initializing S3 client: {e}")
    st.error(f"Unexpected error initializing S3 client: {e}")
    st.stop()

executor = ThreadPoolExecutor(max_workers=MAX_BG_WORKERS, thread_name_prefix="BackgroundTask")

# --- Utility Functions ---

def clean_filename(name: str) -> str:
    """Sanitize a string for safe file or S3 key naming."""
    # Remove invalid characters
    name = re.sub(r'[<>:"/\\|?*\']', '', name)
    # Replace whitespace with underscores
    name = re.sub(r'\s+', '_', name)
    # Limit length (optional, but good practice)
    return name[:100]

def get_task_s3_key() -> str:
    """Get the S3 key for the tasks status file."""
    return f"{S3_TASKS_FOLDER}/tasks.json"

def load_tasks_from_s3() -> Dict[str, dict]:
    """Load background tasks status from S3."""
    s3_key = get_task_s3_key()
    try:
        response = s3_client.get_object(Bucket=S3_BUCKET_NAME, Key=s3_key)
        tasks_data = response['Body'].read().decode('utf-8')
        return json.loads(tasks_data)
    except ClientError as e:
        if e.response['Error']['Code'] == 'NoSuchKey':
            logger.info(f"Tasks file '{s3_key}' not found in S3 bucket '{S3_BUCKET_NAME}'. Starting fresh.")
            return {}
        else:
            logger.error(f"Failed to load tasks file from s3://{S3_BUCKET_NAME}/{s3_key}: {e}")
            return {}
    except (json.JSONDecodeError, Exception) as e:
        logger.error(f"Failed to parse tasks file from s3://{S3_BUCKET_NAME}/{s3_key}: {e}")
        return {} # Return empty dict on error to avoid crashing

def save_tasks_to_s3(tasks: Dict[str, dict]) -> None:
    """Save background tasks status to S3."""
    s3_key = get_task_s3_key()
    try:
        s3_client.put_object(
            Bucket=S3_BUCKET_NAME,
            Key=s3_key,
            Body=json.dumps(tasks, indent=2),
            ContentType='application/json'
        )
        logger.debug(f"Saved tasks status to s3://{S3_BUCKET_NAME}/{s3_key}")
    except ClientError as e:
        logger.error(f"Failed to save tasks file to s3://{S3_BUCKET_NAME}/{s3_key}: {e}")
    except Exception as e:
         logger.error(f"Unexpected error saving tasks file to S3: {e}")

def update_task_status(task_id: str, updates: Dict[str, Any]) -> None:
    """Load, update, and save task status in S3."""
    # Basic retry mechanism for potential S3 eventual consistency or concurrent access issues
    for attempt in range(RETRY_ATTEMPTS):
        tasks = load_tasks_from_s3()
        if task_id in tasks:
            tasks[task_id].update(updates)
            tasks[task_id]["last_updated"] = time.time()
            save_tasks_to_s3(tasks)
            logger.info(f"Updated task {task_id} status: {updates}")
            return
        else:
            logger.warning(f"Task ID {task_id} not found during update attempt {attempt+1}. Retrying...")
            time.sleep(1 * (BACKOFF_FACTOR ** attempt))
    logger.error(f"Failed to find and update task {task_id} after multiple attempts.")


def run_background_task(task_fn, *args, task_keyword: Optional[str] = None, **kwargs) -> str:
    """Run a function in the background, track status via S3."""
    task_id = str(uuid.uuid4())
    # Pass task_id to the background function itself
    future = executor.submit(task_fn, task_id=task_id, *args, **kwargs)

    tasks = load_tasks_from_s3()
    tasks[task_id] = {
        "task_id": task_id,
        "keyword": task_keyword or "N/A",
        "status": "queued",
        "message": "Task submitted to executor.",
        "result": None,
        "error": None,
        "start_time": time.time(),
        "last_updated": time.time()
    }
    save_tasks_to_s3(tasks)
    logger.info(f"Queued background task {task_id} for keyword '{task_keyword}'")

    # Callback to update final status
    def _callback(future_obj):
        final_status = {}
        try:
            result = future_obj.result()
            # The background task itself should update status to 'completed' or 'failed'
            # This callback is a final check / fallback
            tasks = load_tasks_from_s3()
            if tasks.get(task_id, {}).get("status") not in ["completed", "failed"]:
                 final_status = {"status": "completed", "result": result, "message": "Task finished successfully (via callback)."}
                 logger.info(f"Task {task_id} completed (detected by callback). Result: {result}")
            else:
                 logger.info(f"Task {task_id} already finalized (status: {tasks.get(task_id, {}).get('status')}). Callback exiting.")
                 return # Avoid overwriting final status set by the task itself
        except Exception as e:
            logger.error(f"Exception in background task {task_id} (detected by callback): {e}", exc_info=True)
            final_status = {"status": "failed", "error": str(e), "message": "Task failed with exception (via callback)."}

        if final_status:
             update_task_status(task_id, final_status)

    future.add_done_callback(_callback)
    return task_id

# --- S3 Interactions ---

def upload_file_to_s3(local_path: str, s3_bucket: str, s3_key: str) -> bool:
    """Uploads a single file to S3."""
    try:
        s3_client.upload_file(local_path, s3_bucket, s3_key)
        logger.info(f"Successfully uploaded {local_path} to s3://{s3_bucket}/{s3_key}")
        return True
    except ClientError as e:
        logger.error(f"Failed to upload {local_path} to s3://{s3_bucket}/{s3_key}: {e}")
        return False
    except FileNotFoundError:
        logger.error(f"Local file not found for upload: {local_path}")
        return False
    except Exception as e:
        logger.error(f"Unexpected error uploading {local_path} to S3: {e}")
        return False

def upload_directory_to_s3(local_directory: str, s3_bucket: str, s3_prefix: str, delete_local: bool = False) -> bool:
    """Uploads contents of a directory to S3 and optionally removes local files."""
    all_success = True
    local_path_obj = Path(local_directory)
    if not local_path_obj.is_dir():
        logger.error(f"Local directory not found: {local_directory}")
        return False

    for item in local_path_obj.rglob('*'): # Use rglob for recursive
        if item.is_file():
            local_file_path = str(item)
            # Construct S3 key preserving relative structure
            relative_path = item.relative_to(local_path_obj)
            s3_key = f"{s3_prefix}/{relative_path}".replace("\\", "/") # Ensure forward slashes

            if upload_file_to_s3(local_file_path, s3_bucket, s3_key):
                if delete_local:
                    try:
                        os.remove(local_file_path)
                        logger.debug(f"Removed local file: {local_file_path}")
                    except OSError as e:
                        logger.warning(f"Failed to remove local file {local_file_path}: {e}")
                        # Don't mark as failure for the whole upload if deletion fails
            else:
                all_success = False # Mark failure if any upload fails

    # Optionally remove the directory structure if requested and upload was generally successful
    # Be cautious with recursive deletion
    # if delete_local and all_success:
    #     try:
    #         shutil.rmtree(local_directory)
    #         logger.info(f"Removed local directory: {local_directory}")
    #     except OSError as e:
    #         logger.warning(f"Failed to remove local directory {local_directory}: {e}")

    return all_success


# --- YouTube Video Handling ---

def get_youtube_video_links(api_key: str, keyword: str, max_results: int) -> List[Dict[str, str]]:
    """Fetch YouTube video links with rate limit handling."""
    video_links = []
    if not api_key:
        logger.error("YouTube API key is not configured.")
        return []
    try:
        youtube = build("youtube", "v3", developerKey=api_key)
        next_page_token = None
        total_fetched = 0

        while total_fetched < max_results:
            try:
                search_request = youtube.search().list(
                    part="snippet",
                    q=keyword,
                    type="video",
                    maxResults=min(MAX_RESULTS_PER_PAGE, max_results - total_fetched),
                    pageToken=next_page_token
                )
                response = search_request.execute()
                items = response.get("items", [])
                if not items:
                    logger.warning(f"No more video results found for keyword '{keyword}'.")
                    break

                for item in items:
                    if total_fetched >= max_results: break
                    video_id = item.get("id", {}).get("videoId")
                    snippet = item.get("snippet", {})
                    video_title = snippet.get("title", f"Untitled_{video_id}")
                    if video_id:
                        video_url = f"https://www.youtube.com/watch?v={video_id}"
                        # Use clean_filename for the title stored in metadata
                        video_links.append({"Original Title": video_title, "Clean Title": clean_filename(video_title), "Link": video_url})
                        total_fetched += 1

                next_page_token = response.get("nextPageToken")
                if not next_page_token:
                    break # No more pages

            except GoogleHttpError as e:
                if e.resp.status == 403: # Quota Exceeded or similar forbidden errors
                     logger.error(f"YouTube API Forbidden Error (quota likely exceeded): {e}. Stopping search.")
                     st.warning(f"YouTube API Quota likely exceeded. Fetched {total_fetched} videos.")
                     break
                elif e.resp.status == 429: # Rate limit specific
                    logger.warning("YouTube API rate limit hit. Sleeping for 60s.")
                    time.sleep(60) # Simple backoff for rate limits
                else:
                    logger.error(f"YouTube API HTTP error: {e}")
                    raise # Re-raise other HTTP errors
            except Exception as e: # Catch other potential exceptions during API call
                 logger.error(f"Unexpected error during YouTube search API call: {e}")
                 raise # Re-raise

    except Exception as e:
        logger.error(f"Failed to initialize YouTube API or unexpected error: {e}")

    logger.info(f"Fetched {len(video_links)} video links for keyword: {keyword}")
    return video_links

def fetch_youtube_video_data(video_url: str) -> Optional[Dict[str, Any]]:
    """Fetch metadata for a YouTube video using yt-dlp."""
    try:
        ydl_opts = {
            "quiet": True,
            "skip_download": True,
            "extract_flat": False, # Need full metadata
            "forcejson": True, # Force getting JSON metadata
            "socket_timeout": 30, # Add timeout
            }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            # Use cache=False to avoid potential stale metadata? Maybe not necessary.
            info = ydl.extract_info(video_url, download=False)

        if not info:
             logger.warning(f"No metadata extracted for {video_url}")
             return None

        return {
            "original_title": info.get("title", ""),
            "upload_date": info.get("upload_date", ""), # YYYYMMDD format
            "views": info.get("view_count", 0),
            "duration": info.get("duration", 0), # Seconds
            "uploader": info.get("uploader", ""),
            "like_count": info.get("like_count", 0),
            "comment_count": info.get("comment_count", 0), # Often requires login/API, might be inaccurate
            "description": info.get("description", ""),
            "tags": info.get("tags", []),
        }
    except yt_dlp.utils.DownloadError as e:
        # Log specific download errors (like video unavailable, private, etc.)
        logger.error(f"yt-dlp download error fetching metadata for {video_url}: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error fetching metadata for {video_url} with yt-dlp: {e}")
        return None

def download_youtube_video(url: str, output_dir: Path, base_filename: str) -> Optional[Path]:
    """Download a YouTube video in low quality to a specific directory. Returns Path object on success."""
    output_template = output_dir / f"{base_filename}.%(ext)s"
    final_path_mp4 = output_dir / f"{base_filename}.mp4"

    # Check if MP4 already exists
    if final_path_mp4.exists():
        logger.info(f"Video '{final_path_mp4.name}' already exists, skipping download.")
        return final_path_mp4

    try:
        ydl_opts = {
            # Use output template for flexibility with extensions
            "outtmpl": str(output_template),
            # Try to get a reasonable quality single file first, fallback to merge
            "format": "best[height<=480][ext=mp4]/best[ext=mp4]/bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best",
            "merge_output_format": "mp4", # Ensure final is mp4 if merging happens
            "quiet": True,
            "no_warnings": True,
            "socket_timeout": 60, # Longer timeout for download
            "retries": 3, # yt-dlp internal retries
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])

        # Verify the final mp4 file exists after download/merge
        if final_path_mp4.exists():
            logger.info(f"Successfully downloaded '{final_path_mp4.name}'")
            return final_path_mp4
        else:
            # Check if *any* video file was downloaded if the mp4 isn't there
            downloaded_files = list(output_dir.glob(f"{base_filename}.*"))
            if downloaded_files:
                 logger.warning(f"Downloaded video for {url} but final MP4 '{final_path_mp4.name}' not found. Found: {downloaded_files}")
                 # Decide if you want to return the first found file or None
                 # return downloaded_files[0] # Example: return the first match
                 return None # Stricter: only return if mp4 exists
            else:
                 logger.error(f"yt-dlp reported success but no output file found matching '{base_filename}.*' in {output_dir}")
                 return None

    except yt_dlp.utils.DownloadError as e:
        logger.error(f"Download failed for {url}: {e}")
        return None
    except Exception as e:
        logger.error(f"Unexpected error downloading {url}: {e}")
        return None

# --- AI Analysis ---

def wait_for_file_active(file: genai.types.File) -> bool:
    """Wait for Generative AI file to become active."""
    logger.info(f"Waiting for File API file processing: {file.name} ({file.display_name})")
    try:
        file_status = genai.get_file(file.name)
        while file_status.state.name == "PROCESSING":
            logger.debug(f"File {file.name} state: {file_status.state.name}. Waiting...")
            time.sleep(10)
            file_status = genai.get_file(file.name)

        if file_status.state.name == "ACTIVE":
            logger.info(f"File {file.name} is now ACTIVE.")
            return True
        else:
            logger.error(f"File {file.name} processing failed or ended in unexpected state: {file_status.state.name}")
            return False
    except Exception as e:
        logger.error(f"Error waiting for file {file.name}: {e}", exc_info=True)
        return False

def delete_genai_file(file: genai.types.File):
    """Deletes a file uploaded via the GenAI File API."""
    if file:
        try:
            logger.info(f"Attempting to delete GenAI File API file: {file.name}")
            genai.delete_file(file.name)
            logger.info(f"Successfully deleted GenAI File API file: {file.name}")
        except Exception as e:
            logger.warning(f"Failed to delete GenAI File API file {file.name}: {e}")

def analyze_video_with_gemini(video_local_path: Path, retries: int = RETRY_ATTEMPTS) -> Optional[Dict]:
    """Analyze a video using Gemini 1.5 Flash with retry logic and File API cleanup."""
    if not GENAI_API_KEY:
        logger.error("Google Generative AI API key not configured.")
        return None
    if not video_local_path.exists():
        logger.error(f"Video file not found for analysis: {video_local_path}")
        return None

    genai_file = None # Ensure genai_file is defined for finally block
    for attempt in range(retries):
        try:
            logger.info(f"Attempt {attempt + 1}/{retries}: Uploading '{video_local_path.name}' to File API...")
            # Use display_name for better identification in Google's UI
            genai_file = genai.upload_file(str(video_local_path), mime_type="video/mp4", display_name=video_local_path.name)
            if not genai_file:
                raise ValueError("File upload via File API failed, returned None.")
            logger.info(f"File uploaded successfully: Name='{genai_file.name}', URI='{genai_file.uri}'")

            if not wait_for_file_active(genai_file):
                # Don't retry immediately if processing failed, it likely won't succeed without changes
                logger.error(f"File {genai_file.name} processing failed. Aborting analysis for this video.")
                # Deletion happens in finally block
                return None # Indicate failure for this video

            logger.info(f"Attempt {attempt + 1}/{retries}: Analyzing video '{video_local_path.name}' using file {genai_file.name}...")
            model = genai.GenerativeModel(
                # Consider gemini-1.5-pro for potentially better results, but higher cost/latency
                model_name="gemini-1.5-flash-latest",
                generation_config={
                    "temperature": 0.8, # Slightly lower for more factual JSON
                    "top_p": 0.95,
                    "top_k": 64,
                    "max_output_tokens": 8192,
                    "response_mime_type": "application/json"
                },
                 # Add safety settings if needed, e.g., block harmful content
                 # safety_settings=[ ... ]
            )
            # Start chat with the file reference
            chat_session = model.start_chat(history=[{"role": "user", "parts": [genai_file]}])

            # Refined Prompt for more structured JSON
            prompt = (
                "Analyze the provided video thoroughly. Extract the following details and provide the output *strictly* in JSON format, enclosed in a single JSON object. Do not include any text before or after the JSON object. Ensure all string values are properly escaped.\n\n"
                "{\n"
                "    \"analysis_details\": {\n"
                "        \"summary\": \"Provide a concise yet descriptive summary of the video's main content and purpose.\",\n"
                "        \"language\": \"Identify the primary language spoken. Use ISO 639-1 codes (e.g., 'en', 'es') if possible, otherwise the language name.\",\n"
                "        \"duration_seconds\": \"Provide the duration of the video in seconds.\",\n"
                "        \"genre_category\": \"Classify the video's genre (e.g., Tutorial, Advertisement, Vlog, Music Video, News, Comedy Sketch).\",\n"
                "        \"overall_mood\": \"Describe the dominant mood or atmosphere (e.g., Humorous, Suspenseful, Uplifting, Serious, Informative).\",\n"
                "        \"overall_tone\": \"Describe the speaker's or narrator's tone (e.g., Enthusiastic, Formal, Casual, Critical, Persuasive).\",\n"
                "        \"setting_location\": \"Describe the primary physical or virtual setting.\",\n"
                "        \"central_theme\": \"Identify the main underlying theme or message.\",\n"
                "        \"target_audience_profile\": \"Describe the likely intended audience (e.g., Gamers, DIY Enthusiasts, Marketing Professionals, General Public).\",\n"
                "        \"key_objects_present\": [\"List key objects visually present that are relevant to the content.\"],\n"
                "        \"key_topics_discussed\": [\"List the main topics or subjects discussed or shown.\"]\n"
                "    },\n"
                "    \"character_analysis\": {\n"
                "        \"has_people\": \"boolean: Does the video feature identifiable people?\",\n"
                "        \"num_key_characters\": \"integer: Estimate the number of distinct individuals who play a significant role.\",\n"
                "        \"primary_character_persona\": \"Describe the main character's perceived personality or role (if applicable).\",\n"
                "        \"primary_character_emotion\": \"Describe the dominant emotion displayed by the main character (if applicable).\"\n"
                "    },\n"
                "    \"commercial_analysis\": {\n"
                "        \"is_advertisement\": \"boolean: Does the video appear to be primarily an advertisement or promotion?\",\n"
                "        \"brand_mentioned_or_shown\": \"Identify any prominent brand names mentioned or visually displayed.\",\n"
                "        \"product_service_promoted\": \"Identify the specific product or service being promoted (if applicable).\",\n"
                "        \"product_features_highlighted\": [\"List key features or benefits mentioned or shown.\"],\n"
                "        \"call_to_action\": \"Describe any explicit or implicit call to action (e.g., Visit website, Buy now, Subscribe).\"\n"
                "    }\n"
                "}"
            )

            response = chat_session.send_message(prompt)

            # Attempt to parse the JSON response robustly
            try:
                # Remove potential markdown fences (```json ... ```)
                cleaned_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
                analysis_result = json.loads(cleaned_text)
                logger.info(f"Successfully analyzed video '{video_local_path.name}'")
                # Deletion happens in finally block
                return analysis_result
            except json.JSONDecodeError as json_err:
                logger.error(f"Failed to parse JSON response for video '{video_local_path.name}'. Error: {json_err}")
                logger.debug(f"Raw response text: {response.text}")
                # Consider retrying if it was a parsing error, maybe the model output was slightly off
                # If retrying, continue to next iteration
                # If not retrying parse errors, return None or raise
                if attempt < retries - 1:
                     logger.warning("Retrying analysis due to JSON parse error...")
                     time.sleep(SLEEP_TIME_BETWEEN_ANALYSES * (BACKOFF_FACTOR ** attempt))
                     continue # Go to next attempt loop
                else:
                     logger.error("Max retries reached for JSON parsing error.")
                     return None # Failed after retries

        except Exception as e:
            # Check for common API errors (rate limits, quota, etc.)
            error_str = str(e).lower()
            if "rate limit" in error_str or "429" in error_str:
                sleep_time = SLEEP_TIME_BETWEEN_ANALYSES * (BACKOFF_FACTOR ** attempt)
                logger.warning(f"Gemini API rate limit hit (Attempt {attempt + 1}/{retries}). Retrying in {sleep_time:.1f}s: {e}")
                time.sleep(sleep_time)
            elif "quota" in error_str or "403" in error_str:
                 logger.error(f"Gemini API quota exceeded or permission denied. Aborting analysis for '{video_local_path.name}'. Error: {e}")
                 return None # Don't retry quota errors
            elif "500" in error_str or "internal server error" in error_str:
                 sleep_time = SLEEP_TIME_BETWEEN_ANALYSES * (BACKOFF_FACTOR ** attempt)
                 logger.warning(f"Gemini API server error (5xx) (Attempt {attempt + 1}/{retries}). Retrying in {sleep_time:.1f}s: {e}")
                 time.sleep(sleep_time) # Retry server errors
            else:
                logger.error(f"Failed to analyze '{video_local_path.name}' (Attempt {attempt + 1}/{retries}): {e}", exc_info=True)
                # Optional: retry for other transient errors?
                if attempt < retries - 1:
                    sleep_time = SLEEP_TIME_BETWEEN_ANALYSES * (BACKOFF_FACTOR ** attempt)
                    logger.warning(f"Retrying in {sleep_time:.1f}s...")
                    time.sleep(sleep_time)
                else:
                    logger.error(f"Max retries reached for video '{video_local_path.name}'.")
                    return None # Failed after retries
        finally:
            # Ensure the uploaded file is deleted after processing or failure
            if genai_file:
                delete_genai_file(genai_file)
                genai_file = None # Reset after deletion attempt

    logger.error(f"Analysis failed for video '{video_local_path.name}' after {retries} attempts.")
    return None


# --- Data Processing & Background Workflow ---

def flatten_json(data: Dict, parent_key: str = '', sep: str = '_') -> Dict:
    """Flatten nested JSON/Dict for Excel compatibility."""
    items = {}
    for k, v in data.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.update(flatten_json(v, new_key, sep=sep))
        elif isinstance(v, list):
            # Convert list to a string representation (e.g., comma-separated)
            # Handle potential non-string items in list
            items[new_key] = ', '.join(map(str, v))
        else:
            items[new_key] = v
    return items

def merge_dataframes(df_metadata: pd.DataFrame, df_analysis: pd.DataFrame, merge_on_col: str) -> Optional[pd.DataFrame]:
    """Merges two DataFrames, handling potential missing columns"""
    if df_metadata is None or df_analysis is None:
        logger.warning("One or both dataframes for merging are None.")
        return df_metadata if df_analysis is None else df_analysis # Return the non-None one if possible

    if merge_on_col not in df_metadata.columns or merge_on_col not in df_analysis.columns:
        logger.error(f"Merge column '{merge_on_col}' not found in both DataFrames. Cannot merge.")
        # Decide on fallback: return one, or None? Maybe return metadata?
        return df_metadata # Fallback to returning metadata if merge fails

    try:
        # Use outer join to keep all rows, or inner if only fully processed rows are desired
        merged_df = pd.merge(df_metadata, df_analysis, on=merge_on_col, how="left") # Left join preserves all metadata rows
        logger.info(f"Successfully merged dataframes on '{merge_on_col}'. Result shape: {merged_df.shape}")
        return merged_df
    except Exception as e:
        logger.error(f"Error merging dataframes: {e}", exc_info=True)
        return None # Indicate merge failure


def background_process_keyword_search(task_id: str, keyword: str, max_results: int):
    """
    Execute the full pipeline for keyword search in the background.
    Updates task status in S3 via `update_task_status`.
    """
    start_time = time.time()
    project_name = clean_filename(keyword) # Use cleaned keyword as project name
    temp_project_dir = TEMP_PROCESSING_DIR / task_id # Unique temp dir per task
    local_videos_dir = temp_project_dir / "videos"
    local_metadata_file = temp_project_dir / f"{project_name}_metadata.xlsx"
    local_analysis_file = temp_project_dir / f"{project_name}_analysis.xlsx"
    final_output_file_local = temp_project_dir / f"{project_name}_merged_output.xlsx"
    final_output_s3_key = f"{project_name}_merged_output.xlsx" # S3 key for final output

    df_metadata = None
    df_analysis = None

    try:
        # 0. Create temporary directories
        update_task_status(task_id, {"status": "running", "message": "Creating temporary directories..."})
        local_videos_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Task {task_id}: Created temporary directory structure in {temp_project_dir}")

        # 1. Fetch YouTube video links
        update_task_status(task_id, {"message": f"Fetching up to {max_results} video links for '{keyword}'..."})
        video_links_data = get_youtube_video_links(YOUTUBE_API_KEY, keyword, max_results)
        if not video_links_data:
            raise ValueError(f"No video links found for keyword: {keyword}")
        df_links = pd.DataFrame(video_links_data)
        # Use 'Clean Title' for generating filename, 'Link' for processing
        df_links = df_links.rename(columns={"Link": "Youtube_Links"})
        logger.info(f"Task {task_id}: Fetched {len(df_links)} links.")
        update_task_status(task_id, {"message": f"Fetched {len(df_links)} links. Fetching metadata..."})


        # 2. Fetch Metadata for each video
        all_metadata = []
        total_links = len(df_links)
        for index, row in df_links.iterrows():
            url = row["Youtube_Links"]
            clean_title = row["Clean Title"]
            logger.debug(f"Task {task_id}: Fetching metadata for video {index+1}/{total_links}: {url}")
            metadata = fetch_youtube_video_data(url)
            if metadata:
                 # Generate a consistent filename base for download and analysis linking
                 base_filename = f"{clean_title[:50]}_{uuid.uuid4().hex[:8]}" # Add UUID part for uniqueness
                 metadata['base_filename'] = clean_filename(base_filename) # Clean it again just in case
                 metadata['Youtube_Links'] = url # Keep the link
                 metadata['search_keyword'] = keyword # Add keyword context
                 all_metadata.append(metadata)
            else:
                 logger.warning(f"Task {task_id}: Could not fetch metadata for {url}. Skipping.")
            # Update progress intermittently
            if (index + 1) % 10 == 0 or (index + 1) == total_links:
                 update_task_status(task_id, {"message": f"Fetching metadata: {index+1}/{total_links} complete."})
            time.sleep(0.5) # Small delay between metadata fetches


        if not all_metadata:
             raise ValueError("Failed to fetch metadata for any videos.")
        df_metadata = pd.DataFrame(all_metadata)
        # Save intermediate metadata locally (optional but good for debug)
        df_metadata.to_excel(local_metadata_file, index=False)
        logger.info(f"Task {task_id}: Saved metadata for {len(df_metadata)} videos to {local_metadata_file}")
        update_task_status(task_id, {"message": f"Metadata fetched for {len(df_metadata)} videos. Starting downloads..."})


        # 3. Download Videos & Upload to S3 immediately
        downloaded_video_paths = {} # Store mapping: base_filename -> local_path
        s3_video_prefix = f"{S3_VIDEOS_FOLDER}/{project_name}" # S3 path for this project's videos
        total_videos_to_download = len(df_metadata)
        download_success_count = 0

        for index, row in df_metadata.iterrows():
            url = row["Youtube_Links"]
            base_filename = row["base_filename"]
            update_task_status(task_id, {"message": f"Downloading video {index+1}/{total_videos_to_download}: {base_filename}"})
            logger.info(f"Task {task_id}: Downloading video {index+1}/{total_videos_to_download}: {base_filename} from {url}")

            local_video_path = download_youtube_video(url, local_videos_dir, base_filename)

            if local_video_path and local_video_path.exists():
                download_success_count += 1
                downloaded_video_paths[base_filename] = local_video_path
                s3_video_key = f"{s3_video_prefix}/{local_video_path.name}"

                # Upload to S3 immediately after download
                update_task_status(task_id, {"message": f"Uploading video {index+1}/{total_videos_to_download} to S3..."})
                if upload_file_to_s3(str(local_video_path), S3_BUCKET_NAME, s3_video_key):
                    # Optional: Delete local file immediately after successful S3 upload
                    # try:
                    #     os.remove(local_video_path)
                    #     logger.info(f"Task {task_id}: Removed local video {local_video_path.name} after S3 upload.")
                    # except OSError as e:
                    #     logger.warning(f"Task {task_id}: Failed to remove local video {local_video_path.name}: {e}")
                    pass # Keep local for analysis for now
                else:
                    logger.warning(f"Task {task_id}: Failed to upload video {local_video_path.name} to S3. Continuing analysis with local file.")
                    # Keep track of upload failures if needed
            else:
                 logger.warning(f"Task {task_id}: Failed to download video {base_filename} from {url}. Skipping analysis.")
                 # Update metadata df? Maybe add a 'download_status' column
                 df_metadata.loc[index, 'download_status'] = 'failed'


        if download_success_count == 0:
            raise ValueError("Failed to download any videos.")
        logger.info(f"Task {task_id}: Downloaded {download_success_count}/{total_videos_to_download} videos.")
        update_task_status(task_id, {"message": f"Downloaded {download_success_count}/{total_videos_to_download} videos. Starting analysis..."})


        # 4. Analyze Downloaded Videos
        all_analysis_results = []
        total_videos_to_analyze = len(downloaded_video_paths)
        analysis_success_count = 0

        # Ensure GENAI key is available before loop
        if not GENAI_API_KEY:
             raise ValueError("Google Generative AI API key is missing. Cannot perform analysis.")
        genai.configure(api_key=GENAI_API_KEY) # Configure here within the task context

        for i, (base_filename, local_path) in enumerate(downloaded_video_paths.items()):
             update_task_status(task_id, {"message": f"Analyzing video {i+1}/{total_videos_to_analyze}: {base_filename}"})
             logger.info(f"Task {task_id}: Analyzing video {i+1}/{total_videos_to_analyze}: {local_path.name}")

             analysis_result = analyze_video_with_gemini(local_path)

             if analysis_result:
                 analysis_success_count += 1
                 # Flatten the result and add the filename for merging
                 flat_analysis = flatten_json(analysis_result)
                 flat_analysis['base_filename'] = base_filename # Use the same key for merging
                 all_analysis_results.append(flat_analysis)
                 # Optional: Add a small delay even after successful analysis if hitting limits often
                 time.sleep(5) # Reduced delay between analyses
             else:
                 logger.warning(f"Task {task_id}: Failed to analyze video {local_path.name}.")
                 # Optionally add failed status to metadata df
                 idx = df_metadata.index[df_metadata['base_filename'] == base_filename].tolist()
                 if idx:
                     df_metadata.loc[idx[0], 'analysis_status'] = 'failed'

             # Sleep regardless of success/failure if needed to respect potential API limits further
             # time.sleep(SLEEP_TIME_BETWEEN_ANALYSES) # Use constant


        if analysis_success_count == 0:
             logger.warning(f"Task {task_id}: Analysis failed for all downloaded videos.")
             # Proceed to merge metadata only if desired, or raise error
             # raise ValueError("Failed to analyze any videos.")
             df_analysis = pd.DataFrame() # Create empty df if no analysis succeeded
        else:
             df_analysis = pd.DataFrame(all_analysis_results)
             df_analysis.to_excel(local_analysis_file, index=False) # Save intermediate analysis
             logger.info(f"Task {task_id}: Analyzed {analysis_success_count}/{total_videos_to_analyze} videos. Saved analysis to {local_analysis_file}")

        update_task_status(task_id, {"message": f"Analysis complete for {analysis_success_count}/{total_videos_to_analyze} videos. Merging data..."})


        # 5. Merge Metadata and Analysis Data
        logger.info(f"Task {task_id}: Merging metadata and analysis data...")
        final_df = merge_dataframes(df_metadata, df_analysis, merge_on_col='base_filename')

        if final_df is None or final_df.empty:
            # Handle cases where merge fails or results in empty dataframe
            if df_metadata is not None and not df_metadata.empty:
                 logger.warning(f"Task {task_id}: Merge failed or resulted in empty data. Saving metadata only.")
                 final_df = df_metadata # Fallback to metadata
            else:
                 raise ValueError("Merging failed and no metadata available.")

        final_df.to_excel(final_output_file_local, index=False)
        logger.info(f"Task {task_id}: Merged data saved locally to {final_output_file_local}")


        # 6. Upload Final Output to S3 Output Bucket
        update_task_status(task_id, {"message": "Uploading final merged output file..."})
        if upload_file_to_s3(str(final_output_file_local), S3_OUTPUT_BUCKET_NAME, final_output_s3_key):
            logger.info(f"Task {task_id}: Successfully uploaded final output to s3://{S3_OUTPUT_BUCKET_NAME}/{final_output_s3_key}")
            final_result_location = f"s3://{S3_OUTPUT_BUCKET_NAME}/{final_output_s3_key}"
            update_task_status(task_id, {"status": "completed", "message": "Processing complete.", "result": final_result_location})
            return final_result_location
        else:
            raise IOError(f"Failed to upload final output file to S3.")


    except Exception as e:
        logger.error(f"Error in background task {task_id} for keyword '{keyword}': {e}", exc_info=True)
        update_task_status(task_id, {"status": "failed", "error": str(e), "message": f"Task failed: {e}"})
        return None # Indicate failure
    finally:
        # 7. Cleanup Temporary Files/Directory
        try:
            if TEMP_PROCESSING_DIR.exists() and task_id in str(temp_project_dir): # Safety check
                shutil.rmtree(temp_project_dir)
                logger.info(f"Task {task_id}: Cleaned up temporary directory {temp_project_dir}")
        except Exception as clean_e:
            logger.warning(f"Task {task_id}: Failed to cleanup temporary directory {temp_project_dir}: {clean_e}")
        end_time = time.time()
        logger.info(f"Task {task_id} finished in {end_time - start_time:.2f} seconds.")


# --- Streamlit App ---

def list_output_files() -> List[Tuple[str, str]]:
    """List files in the S3 output bucket and generate presigned URLs for download."""
    output_files = []
    try:
        response = s3_client.list_objects_v2(Bucket=S3_OUTPUT_BUCKET_NAME)
        if 'Contents' in response:
            for obj in response['Contents']:
                key = obj['Key']
                # Generate presigned URL (expires in 1 hour by default)
                try:
                    url = s3_client.generate_presigned_url('get_object',
                                                          Params={'Bucket': S3_OUTPUT_BUCKET_NAME, 'Key': key},
                                                          ExpiresIn=3600) # 1 hour expiry
                    output_files.append((key, url))
                except ClientError as url_e:
                     logger.error(f"Failed to generate presigned URL for {key}: {url_e}")
                     output_files.append((key, "#error")) # Indicate error generating URL
        return output_files
    except ClientError as e:
        logger.error(f"Failed to list S3 output files from bucket '{S3_OUTPUT_BUCKET_NAME}': {e}")
        st.error(f"Failed to list S3 output files: {e}")
        return []

def app():
    """Main Streamlit application."""
    # Initial checks for essential configurations
    if not all([YOUTUBE_API_KEY, GENAI_API_KEY, S3_BUCKET_NAME, S3_OUTPUT_BUCKET_NAME]):
        st.error("ERROR: Critical environment variables (YOUTUBE_API_KEY, GENAI_API_KEY, AWS_BUCKET_NAME, S3_OUTPUT_BUCKET_NAME) are missing! Please configure them.")
        st.stop() # Stop if critical env vars are missing


    st.set_page_config(layout="wide") # Use wider layout

    st.sidebar.title("🎬 YouTube Video Analyzer")
    option = st.sidebar.radio(
        "Select Mode",
        ["Keyword Search (Background)", "Background Task Status", "View Output Files"], # Simplified modes
        help="Choose how to process videos or view results."
    )

    # --- Keyword Search Mode ---
    if option == "Keyword Search (Background)":
        st.header("🚀 Keyword Search (Background Processing)")
        st.markdown("Enter a keyword to find YouTube videos. The system will fetch links, download videos, analyze them using AI, and store the results. This process runs in the background.")

        keyword = st.text_input("Enter Search Keyword:", placeholder="e.g., product reviews, tech tutorials")
        max_results = st.slider("Max Videos to Process:", min_value=1, max_value=50, value=5, step=1, help="Number of video links to fetch. More videos take longer.") # Default to lower value

        if st.button("Start Background Job", type="primary"):
            if keyword:
                with st.spinner(f"Submitting job for keyword: '{keyword}'..."):
                    # Sanitize keyword for use in filenames/paths if needed, although background_process handles it
                    clean_kw = clean_filename(keyword)
                    task_id = run_background_task(
                        background_process_keyword_search,
                        keyword=keyword,
                        max_results=max_results,
                        task_keyword=keyword # Pass keyword for display in status
                        )
                st.success(f"✅ Background task submitted successfully! Task ID: `{task_id}`")
                st.info("You can monitor the progress in the 'Background Task Status' tab.")
                st.balloons()
            else:
                st.warning("⚠️ Please enter a keyword to search.")

    # --- Background Task Status Mode ---
    elif option == "Background Task Status":
        st.header("📊 Background Task Status")
        st.markdown("Monitor the progress of submitted background jobs.")

        if st.button("Refresh Status"):
            st.rerun() # Simple way to refresh

        tasks = load_tasks_from_s3()
        if not tasks:
            st.info("No background tasks found.")
        else:
            st.write(f"Found {len(tasks)} tasks. Displaying latest first:")
            # Sort tasks by start time, newest first
            sorted_tasks = sorted(tasks.items(), key=lambda item: item[1].get('start_time', 0), reverse=True)

            for task_id, task_info in sorted_tasks:
                status = task_info.get('status', 'unknown')
                keyword = task_info.get('keyword', 'N/A')
                message = task_info.get('message', 'No message.')
                start_time_ts = task_info.get('start_time')
                last_updated_ts = task_info.get('last_updated')

                start_time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time_ts)) if start_time_ts else "N/A"
                last_updated_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(last_updated_ts)) if last_updated_ts else "N/A"


                with st.expander(f"Task ID: `{task_id}` | Keyword: `{keyword}` | Status: `{status.upper()}`"):
                    st.write(f"**Status:** {status.upper()}")
                    st.write(f"**Keyword:** {keyword}")
                    st.write(f"**Current Message:** {message}")
                    st.write(f"**Started:** {start_time_str}")
                    st.write(f"**Last Update:** {last_updated_str}")

                    if status == "completed":
                        st.success("✅ Task completed successfully.")
                        result_location = task_info.get("result")
                        if result_location and result_location.startswith("s3://"):
                             st.info(f"Output file located at: {result_location}")
                             # Optional: Add a direct download link if possible (needs presigned URL logic here too)
                             try:
                                 bucket, key = result_location.replace("s3://", "").split("/", 1)
                                 url = s3_client.generate_presigned_url('get_object', Params={'Bucket': bucket, 'Key': key}, ExpiresIn=3600)
                                 st.markdown(f"[Download Output File]({url})", unsafe_allow_html=True)
                             except Exception as url_e:
                                 st.warning(f"Could not generate download link: {url_e}")

                    elif status == "failed":
                        st.error(f"❌ Task failed.")
                        error_msg = task_info.get("error", "No specific error message available.")
                        st.code(error_msg, language=None) # Display error in code block
                    elif status == "running" or status == "queued":
                        st.info(f"⏳ Task is currently {status}.")
                        # Add a simple progress indicator if possible (e.g., based on message)
                        if "Downloading video" in message or "Analyzing video" in message:
                             # Basic progress text parsing
                             match = re.search(r'(\d+)/(\d+)', message)
                             if match:
                                 current, total = map(int, match.groups())
                                 progress_percent = int((current / total) * 100) if total > 0 else 0
                                 st.progress(progress_percent, text=f"{message} ({progress_percent}%)")
                             else:
                                 st.progress(50, text=message) # Indeterminate progress if parsing fails
                        elif "Fetching metadata" in message or "Merging data" in message:
                             st.progress(25, text=message) # Arbitrary progress for these stages


    # --- View Output Files Mode ---
    elif option == "View Output Files":
        st.header("📂 View Output Files")
        st.markdown(f"Listing finalized output files stored in the S3 bucket: `{S3_OUTPUT_BUCKET_NAME}`.")

        if st.button("Refresh File List"):
            st.rerun()

        output_files = list_output_files()

        if not output_files:
            st.info("No output files found in the S3 bucket.")
        else:
            st.write(f"Found {len(output_files)} files:")
            df_files = pd.DataFrame(output_files, columns=["Filename", "Download URL"])

            # Display as a table with clickable links
            st.dataframe(
                 df_files,
                 column_config={
                      "Filename": st.column_config.TextColumn("S3 File Key"),
                      "Download URL": st.column_config.LinkColumn(
                           "Download Link",
                           help="Click to download the file (link expires in 1 hour)",
                           display_text="Download"
                      )
                 },
                 hide_index=True,
                 use_container_width=True
            )


    # --- Removed Modes (Single Video / File Mode) ---
    # These modes performed processing synchronously in the foreground,
    # which is less suitable for a t2.micro. They could be adapted
    # to also use the background task system if needed.
    # For now, they are removed for simplicity and focus on background processing.

# --- Main Execution Guard ---
if __name__ == "__main__":
    # Ensure temp directory exists
    TEMP_PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    app()