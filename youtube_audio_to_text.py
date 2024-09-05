import whisper
import googleapiclient.discovery
from google.oauth2 import service_account
import yt_dlp as youtube_dl
import uuid
import os

def transcribe_youtube_video(url, api_key):
    """Transcribes a YouTube video using Whisper and YouTube Data API metadata.

    Args:
        url: The URL of the YouTube video.
        api_key: Your YouTube Data API key.

    Returns:
        A dictionary containing transcribed text, title, description, published date,
        and a flag indicating if captions were found.
    """

    # Create a YouTube Data API client
    youtube = googleapiclient.discovery.build("youtube", "v3", developerKey=api_key)

    # Extract the video ID from the URL
    video_id = url.split("v=")[1]

    # Get video metadata
    video_response = youtube.videos().list(
        part="snippet,contentDetails",
        id=video_id
    ).execute()

    try:
        video_data = video_response['items'][0]
        title = video_data['snippet']['title']
        description = video_data['snippet']['description']
        published_at = video_data['snippet']['publishedAt']
    except IndexError:
        print("Error: Video not found.")
        return {}

    # Check for captions
    captions_response = youtube.captions().list(
        part="snippet",
        videoId=video_id
    ).execute()

    has_captions = bool(captions_response.get('items'))

    # Download the audio using yt-dlp
    audio_file_name = f"video_{video_id}.mp3"
    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': audio_file_name,
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
    }

    try:
        with youtube_dl.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
    except Exception as e:
        print(f"Error downloading audio: {e}")
        return {}

    # Verify the audio file
    audio_file_path = f"{audio_file_name}.mp3"
    if not os.path.exists(audio_file_path):
        print("Error: Audio file not found.")
        return {}

    # Transcribe using Whisper
    try:
        transcribed_text = transcribe_audio_with_whisper(audio_file_path)
    except Exception as e:
        print(f"Error transcribing audio: {e}")
        return {}

    return {
        "transcribed_text": transcribed_text,
        "title": title,
        "description": description,
        "published_at": published_at,
        "has_captions": has_captions
    }

def transcribe_audio_with_whisper(audio_file):
    """Transcribes an audio file using Whisper.

    Args:
        audio_file: The path to the audio file.

    Returns:
        The transcribed text.
    """

    model = whisper.load_model("base")
    result = model.transcribe(audio_file)
    return result["text"]

# Replace 'YOUR_API_KEY' with your actual API key
api_key = 'YOUR_API_KEY'

# Get the YouTube video URL from the user
url = input("Enter the YouTube video URL: ")

# Transcribe the video
result = transcribe_youtube_video(url, api_key)

if result:
    print(f"Title: {result['title']}")
    print(f"Description: {result['description']}")
    print(f"Published at: {result['published_at']}")
    print(f"Has captions: {result['has_captions']}")
    print(f"Transcribed text: {result['transcribed_text']}")
else:
    print("Transcription failed. Please check the YouTube video URL and try again.")