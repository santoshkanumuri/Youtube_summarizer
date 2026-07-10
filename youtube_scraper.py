from googleapiclient.discovery import build
import openpyxl
import os

def get_youtube_video_details(api_key, keyword, max_results=100):
    youtube = build('youtube', 'v3', developerKey=api_key)

    # Search for videos
    search_request = youtube.search().list(
        part='snippet',
        q=keyword,
        type='video',
        maxResults=max_results
    )

    search_response = search_request.execute()

    video_details = []
    for index, item in enumerate(search_response['items'], start=1):
        video_id = item['id']['videoId']
        video_title = item['snippet']['title']
        video_url = f"https://www.youtube.com/watch?v={video_id}"

        # Get video statistics
        stats_request = youtube.videos().list(
            part='statistics',
            id=video_id
        )
        stats_response = stats_request.execute()

        like_count = stats_response['items'][0]['statistics'].get('likeCount', '0')

        # Get all comments
        comments = []
        try:
            comments_request = youtube.commentThreads().list(
                part='snippet',
                videoId=video_id,
                maxResults=100
            )

            while comments_request:
                comments_response = comments_request.execute()
                for comment_item in comments_response['items']:
                    comment = comment_item['snippet']['topLevelComment']['snippet']['textDisplay']
                    comments.append(comment)
                comments_request = youtube.commentThreads().list_next(comments_request, comments_response)
        except Exception as e:
            print(f"Comments are disabled for video: {video_title}")

        # Save comments to a new Excel file
        comments_filename = f"./comments/{index}.xlsx"
        save_comments_to_excel(comments, comments_filename)

        video_details.append({
            'SerialNo': index,
            'Title': video_title,
            'CommentsFilePath': os.path.abspath(comments_filename)
        })

    return video_details


def save_comments_to_excel(comments, filename):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Comments'

    # Create headers
    headers = ['Comment']
    sheet.append(headers)

    # Add comments
    for comment in comments:
        sheet.append([comment])

    # Adjust column width
    sheet.column_dimensions['A'].width = 100

    # Save the workbook
    workbook.save(filename)


def save_summary_to_excel(video_details, filename='youtube_summary.xlsx'):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'YouTube Video Summary'

    # Create headers
    headers = ['SerialNo', 'Title', 'CommentsFilePath']
    sheet.append(headers)

    # Add video details
    for video in video_details:
        sheet.append([video['SerialNo'], video['Title'], video['CommentsFilePath']])

    # Adjust column widths
    for column in sheet.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width

    # Save the workbook
    workbook.save(filename)


# Example usage
api_key = os.getenv('YOUTUBE_API_KEY') or os.getenv('GOOGLE_API_KEY')
if not api_key:
    raise RuntimeError("Set YOUTUBE_API_KEY or GOOGLE_API_KEY")
keyword = 'menopause'
video_details = get_youtube_video_details(api_key, keyword)
save_summary_to_excel(video_details)

print("Video details and comments saved to Excel.")
