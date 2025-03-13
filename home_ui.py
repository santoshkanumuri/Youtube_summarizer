import os.path
import streamlit as st
from video_links_data import save_video_links_to_excel
from process_excel_links import (fetch_youtube_video_data,download_youtube_video,extract_youtube_links,
                                 update_excel_with_video_data,downloader,merge_data)
import pandas as pd
from summarizer import analyze_video,analyze_folder


def keyword_search(keyword,time_limit,video_limit):
    # api_key = '***REMOVED***'
    # api_key="***REMOVED***"
    api_key=""
    st.write(video_limit)
    save_video_links_to_excel(api_key, keyword, int(video_limit))


def clean_title(title):
    return "".join([c for c in title if c.isalpha() or c.isdigit() or c == " "]).rstrip()

def app():

    option=st.sidebar.radio('Select an option', ['Keyword Search', 'Single Video', 'File Mode'])
    if option=='Keyword Search':
        st.title('Youtube video summarizer - Keyword Search')
        st.write('Keyword Search Mode')
        keyword=st.text_input('Enter the keyword to search')
        time_limit=st.slider('Select the time limit for the video in seconds.', min_value=1, max_value=1000)
        video_limit=st.slider('Select the number of videos to search', min_value=1, max_value=300)
        if st.button('Search'):
            keyword_search(keyword,time_limit,video_limit)

    elif option == 'Single Video':
        st.title('Youtube video summarizer - Single video')
        st.write('Single Video Mode')
        url=st.text_input('Enter the video link')
        video_path = './videos/single_url/'
        if st.button('Summarize'):
            with st.status("Processing..."):
                st.write('Fetching video data...')
                data=fetch_youtube_video_data(url)
                file_name=clean_title(data['title'])[:30]
                st.write('Downloading video...')
                downloaded=download_youtube_video(url, video_path, file_name[:30])
                video_path=video_path+file_name+'.mp4'
                if os.path.exists(video_path):
                    downloaded=True
                if downloaded:
                    st.write('video saved to ',video_path)
                    try:
                        st.write('Analyzing the video...')
                        output=analyze_video(video_path)
                    except Exception as e:
                        st.error(f"An error occurred: {e}")
            st.table(output)

    elif option == 'File Mode':
        st.title('Youtube video summarizer - File Mode')
        st.write('File Mode')
        file=st.text_input('Enter the file path')
        project_name=st.text_input('Enter the project folder name')
        if(st.button('Extract Links')):
            if file is not None:
                st.write('Extracting youtube links from the file...')
                excel_file=extract_youtube_links(file,project_name)
                st.write('Fetching video data from excel file links...')
                excel_file=update_excel_with_video_data(excel_file)
                st.write('Downloading videos...')
                downloader(excel_file,project_name)
                st.write('All videos downloaded.')
                videos_folder_path = f'./videos/{project_name}/'
                st.write('Analyzing videos...')
                summary_file=analyze_folder(videos_folder_path, project_name)
                st.write('All videos analyzed, saved to ',summary_file)
                st.write('Merging data...')
                merge_data(excel_file,summary_file,project_name)
                st.write('All data merged and saved to output folder.')







            else:
                st.write('Please enter valid file path')



if __name__ == '__main__':
    app()