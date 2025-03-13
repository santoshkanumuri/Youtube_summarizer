import pandas as pd

# Load the two Excel sheets
df1 = pd.read_excel('youtube_video_links.xlsx')  # Replace with your actual file name
df2 = pd.read_excel('video_analysis_results1.xlsx')  # Replace with your actual file name

def clean_title(title):
    return "".join([c for c in title if c.isalpha() or c.isdigit() or c==' ' or c=="."]).rstrip()

# Clean the 'Title' column in both DataFrames
df1['file_name'] = df1['file_name'].apply(clean_title)
df2['file_name'] = df2['file_name'].apply(clean_title)

merged_data = pd.merge(df1, df2, how='inner', on='file_name')

# Save the merged DataFrame to a new Excel file
merged_data.to_excel('merged_videos_data.xlsx', index=False)

print("Merge complete. Saved as 'merged_videos_data.xlsx'.")
