import streamlit as st
import subprocess
import os
from youtube_transcript_api import YouTubeTranscriptApi
from pytube import YouTube
import youtube_dl

def run_command(youtube_url, output_dir):
   
    # Clean up the YouTube URL. It may have extra characters at the end for the timestamp
    # Remove time marker from the YouTube URL
    youtube_url = youtube_url.split('&')[0]

    # Create a YouTube object
    yt = YouTube(youtube_url)

    # Get the video title
    video_title = yt.title

    # Extract video transcript
    transcript = YouTubeTranscriptApi.get_transcript(youtube_url.split('=')[-1])

    # Convert transcript to text
    transcript_text = ' '.join([x['text'] for x in transcript])

    # Save full transcript to a file
    with open(f"{output_dir}/{video_title}_full_transcript.txt", 'w') as f:
        f.write(youtube_url + "\n" + transcript_text)

    # Command to extract transcript and process it using fabric
    command = f'echo "{transcript_text}" | fabric --pattern extract_wisdom > "{output_dir}/{video_title}.md"'
    
    # Run the command in the shell
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    
    if result.returncode == 0:
        st.success(f"File saved to: {output_dir}/{video_title}.md")
    else:
        st.error(f"Error: {result.stderr}")

    return video_title

def main():
    st.title("YouTube Wisdom Extractor")
    
    # Input fields for YouTube URL and output directory
    youtube_url = st.text_input("Enter YouTube URL")
    output_dir = st.text_input("Specify output directory", os.path.expanduser("/Users/gerard/Library/CloudStorage/OneDrive-Personal/3. Resources/YouTube_Notes"))

    # Button to trigger the process
    if st.button("Extract Wisdom"):
        if youtube_url and output_dir:
            run_command(youtube_url, output_dir)
        else:
            st.error("Please provide both YouTube URL and output directory")

if __name__ == "__main__":
    main()
