
import os
import time
import google.generativeai as genai

genai.configure(api_key="***REMOVED***")

def upload_to_gemini(path, mime_type=None):
  """Uploads the given file to Gemini.

  See https://ai.google.dev/gemini-api/docs/prompting_with_media
  """
  file = genai.upload_file(path, mime_type=mime_type)
  print(f"Uploaded file '{file.display_name}' as: {file.uri}")
  return file

def wait_for_files_active(files):
  """Waits for the given files to be active.

  Some files uploaded to the Gemini API need to be processed before they can be
  used as prompt inputs. The status can be seen by querying the file's "state"
  field.

  This implementation uses a simple blocking polling loop. Production code
  should probably employ a more sophisticated approach.
  """
  print("Waiting for file processing...")
  for name in (file.name for file in files):
    file = genai.get_file(name)
    while file.state.name == "PROCESSING":
      print(".", end="", flush=True)
      time.sleep(30)
      file = genai.get_file(name)
    if file.state.name != "ACTIVE":
      raise Exception(f"File {file.name} failed to process")
  print("...all files ready")

# Create the model
generation_config = {
  "temperature": 1,
  "top_p": 0.95,
  "top_k": 64,
  "max_output_tokens": 8192,
  "response_mime_type": "text/plain",
}
print("Creating model...")
model = genai.GenerativeModel(
  model_name="gemini-1.5-flash",
  generation_config=generation_config,
  # safety_settings = Adjust safety settings
  # See https://ai.google.dev/gemini-api/docs/safety-settings
)

# TODO Make these files available on the local file system
# You may need to update the file paths

# Upload the video file
print("Uploading video file...")
files = [upload_to_gemini("C:\\Users\\Pavan\\PycharmProjects\\youtube_data\\videos\\Menopil Plus Menopause Support.mp4", mime_type="video/mp4")]


# Some files have a processing delay. Wait for them to be ready.
print("Waiting for files to be ready...")
wait_for_files_active(files)

print("Starting chat session...")
chat_session = model.start_chat(
  history=[{
      "role": "user",
      "parts": [
        files[0],
      ],
    },
  ]
)

print("Chat session started.")
response = chat_session.send_message(f'''For the provided video above, analyze the video and provide me the following details in json format output."details": 
      "video_summary": "Provide a clear descriptive summary of the video's content.",
      "number_of_characters": "Count the number of distinct characters or people appearing in the video that have impact on the video.",
      "video_total_emotion": "Analyze the overall emotional tone of the video (e.g., positive, negative, neutral).",
      "video_duration": "Provide the duration of the video in seconds.",
        "video_language": "Identify the primary language spoken in the video.",
        "video_genre": "Identify the genre or category of the video (e.g., comedy, drama, action).",
        "video_mood": "Identify the mood or atmosphere of the video (e.g., suspenseful, romantic, humorous).",
        "video_tone": "Identify the tone or style of the video (e.g., serious, light-hearted, satirical).",
        "color_palette": "Identify the primary color palette used in the video.",
        "words_used": "Identify any key words or phrases that are repeated or emphasized in the video.",
      "brand_name": "Identify the brand name associated with the video, especially if it's an advertisement.",
        "product_name": "Identify the product or service being promoted in the video.",
        "product_features": "Identify the key features or benefits of the product or service.",
        "product_target_audience": "Identify the target audience or demographic for the product or service.",
        "product_message": "Identify the central message or value proposition of the product or service.",
    ''')

print(response.text)