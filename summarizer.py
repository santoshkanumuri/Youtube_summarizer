import os
import time
import google.generativeai as genai
import pandas as pd
import json
import dotenv


# Configure the API key
genai.configure(api_key="***REMOVED***")


def upload_to_gemini(path, mime_type=None):
    """Uploads the given file to Gemini."""
    file = genai.upload_file(path, mime_type=mime_type)
    print(f"Uploaded file '{file.display_name}' as: {file.uri}")
    return file


def wait_for_file_active(file):
    """Waits for the given file to be active."""
    print("Waiting for file processing...")
    file_status = genai.get_file(file.name)
    while file_status.state.name == "PROCESSING":
        print(".", end="", flush=True)
        time.sleep(10)
        file_status = genai.get_file(file.name)
    if file_status.state.name != "ACTIVE":
        raise Exception(f"File {file.name} failed to process")
    print("...file ready")
    return file_status


def analyze_video(video_path):
    print(f"Uploading video file: {video_path}")
    file = upload_to_gemini(video_path, mime_type="video/mp4")

    print("Waiting for file to be ready...")
    file_status = wait_for_file_active(file)

    print("Starting chat session...")
    chat_session = genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        generation_config={
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 64,
            "max_output_tokens": 8192,
            "response_mime_type": "application/json",
        },
    ).start_chat(history=[{
      "role": "user",
      "parts": [
        file,
      ],
    },])

    response = chat_session.send_message('''For the given video, analyze the video and provide me the following details in JSON format:
    "details": {
        "video_summary": "Provide a clear descriptive summary of the video's content.",
        "number_of_characters": "Count the number of distinct characters or people appearing in the video that have impact on the video.",
        "video_total_emotion": "Analyze the overall emotional tone of the video (e.g., positive, negative, neutral).",
        "video_duration": "Provide the duration of the video in seconds.",
        "video_language": "Identify the primary language spoken in the video.",
        "video_genre": "Identify the genre or category of the video (e.g., comedy, drama, action).",
        "video_mood": "Identify the mood or atmosphere of the video (e.g., suspenseful, romantic, humorous).",
        "video_tone": "Identify the tone or style of the video (e.g., serious, light-hearted, satirical).",
        "video_setting": "Identify the primary setting or location of the video.",
        "video_theme": "Identify the central theme or message of the video.",
        "video_impact": "Analyze the impact or significance of the video on the viewer.",
        "video_target_audience": "Identify the target audience or demographic for the video.",
        "video_influences": "Identify any notable influences or inspirations for the video.",
        "video_references": "Identify any references or allusions made in the video.",
        "color_palette": "Identify the primary color palette used in the video.",
        "words_used": "Identify any key words or phrases that are repeated or emphasized in the video.",
        "brand_name": "Identify the brand name associated with the video, especially if it's an advertisement.",
        "product_name": "Identify the product or service being promoted in the video.",
        "product_features": "Identify the key features or benefits of the product or service.",
        "product_target_audience": "Identify the target audience or demographic for the product or service.",
        "product_message": "Identify the central message or value proposition of the product or service.",
        "product_call_to_action": "Identify the call-to-action or desired response from the viewer.",
        "Personality Trait": "Identify the personality of the person in the video. Identify in these categories:  Extraversion, Agreeableness, Neuroticism, Conscientiousness, Openness to Experience, Narcissism, Machiavellianism, Psychopathy, Maverickism",
        "Terms Associated": "Identify the terms associated with personality traits",
        "This person will be": "Identify the personality of the person in the video Example-(Will demonstrate energy, action, assertiveness in his words and actions), (Will act whimsically, without any consideration of surrounding or others. They have a superficial charm that lures people, but they lack emotions or empathy)",
        "Person Tone": "Identify the tone of the person in the video. Example-(Serious, Light-hearted, Satirical)",
        "Person Mood": "Identify the mood of the person in the video. Example-(Suspenseful, Romantic, Humorous)",
        "Person emotion": "Identify the emotion of the person in the video. Example-(Positive, Negative, Neutral)"
    }''')
    data=json.loads(response.text)
    return data['details']


def analyze_folder(folder_path, project_name):
    output_excel=f'./output_files/{project_name}/processing/ai_analyzed_video_data.xlsx'
    if not os.path.exists(os.path.dirname(output_excel)):
        os.makedirs(os.path.dirname(output_excel))
    video_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if
                   f.endswith('.mp4')]
    all_data = []

    for video in video_files:
        try:
            print(f"Processing video: {video}")
            video_data = analyze_video(video)
            print(f"Analysis complete for {video}")
            time.sleep(60)
            video_data['File_Name'] = os.path.basename(video)
            all_data.append(video_data)
        except Exception as e:
            print(f"Failed to process video {video}: {e}")



    # Convert the collected data into a DataFrame and save it as an Excel file
    df = pd.DataFrame(all_data)
    df.to_excel(output_excel, index=False)
    print(f"Analysis complete. Results saved to {output_excel}")
    return output_excel


if __name__ == '__main__':
    pass
