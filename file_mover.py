import os
import shutil
import pandas as pd

# Specify the paths
excel_file_path = "C:\\Users\\Pavan\\PycharmProjects\\youtube_data\\video_analysis_results.xlsx"  # Replace with the path to your Excel file
source_folder = "C:\\Users\\Pavan\\PycharmProjects\\youtube_data\\videos"  # Replace with the path to the folder containing your videos
completed_folder = "C:\\Users\\Pavan\\PycharmProjects\\youtube_data\\completed"

# Create the completed folder if it doesn't exist
os.makedirs(completed_folder, exist_ok=True)

# Read the Excel file
df = pd.read_excel(excel_file_path)

# Check if the file_name column exists
if 'file_name' not in df.columns:
    print("The 'file_name' column does not exist in the Excel file.")
else:
    # Iterate through each file name in the 'file_name' column
    for file_name in df['file_name']:
        # Construct the full file path
        source_file_path = os.path.join(source_folder, file_name)

        # Check if the file exists
        if os.path.exists(source_file_path):
            # Move the file to the completed folder
            shutil.move(source_file_path, completed_folder)
            print(f"Moved: {file_name}")
        else:
            print(f"File not found: {file_name}")

print("File moving process completed.")
