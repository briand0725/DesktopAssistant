import subprocess

import nltk

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')

import time
import speech_recognition as sr
from nltk import word_tokenize, pos_tag
import webbrowser
import datetime
from AppOpener import open, close
import os
import difflib
import imaplib
import email
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from fuzzywuzzy import process



# Specify the directory path where your Word documents are stored
directory_path = r"C:\Users\wilso\OneDrive"

# Get a list of all files in the directory
all_files = os.listdir(directory_path)

# Filter only Word documents
word_docs = []
for file in all_files:
    if file.lower().endswith(".docx"):
        word_docs.append(file)

# Now, word_docs contains the titles of all Word documents in the specified directory
existing_word_docs = [doc[:-5] for doc in word_docs]  # Remove the ".docx" extension
print(existing_word_docs)

today = datetime.datetime.today()

class_schedule = {
    "monday": ["9:05 am : Python at GOODW 190"],
    "tuesday": ["8:00 am : CSP at TORG 1060", "11:00 am : Sociology at MCB 100",
                "12:30 pm : Psych at MCB 100", "3:30 pm : Linear at MCB 230", "5:00 pm : Engineering at GOODW 135"],
    "wednesday": ["9:05 am : Python at GOODW 190"],
    "thursday": ["8:00 am : CSP at TORG 1060", "11:00 am : Sociology at MCB 100",
                "12:30 pm : Psych at MCB 100", "3:30 pm : Linear at MCB 230", "5:00 pm : Engineering at GOODW 135"],
    "friday": ["9:05 am : Python at GOODW 190"]
}

def open_kindle():
    open("kindle", match_closest=True)
    print("kindle has been opened")


def close_kindle():
    close("kindle", match_closest=True)
    print("kindle has been closed")


def open_youtube():  # opens youtube
    webbrowser.open("https://www.youtube.com/")
    print("youtube has been opened")


def open_discord():
    open("discord", match_closest= True)
    print("discord has been opened")


def close_discord():
    close("discord", match_closest= True)
    print("discord has been closed")


def open_spotify():
    open("spotify", match_closest= True)
    print("spotify has been opened")


def close_spotify():
    close("spotify", match_closest= True)
    print("spotify has been closed")


def open_gmail():
    webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
    print("gmail has been opened")


def open_outlook():
    webbrowser.open("https://outlook.office.com/mail/")
    print("outlook has been opened")


def open_canvas():
    webbrowser.open("https://canvas.vt.edu/")
    print("canvas has been opened")


def check_weather():
    pass


def schedule():
    schedule_info = ""
    if today.weekday() == 0:
        schedule_info = class_schedule["monday"]
    elif today.weekday() == 1:
        schedule_info = class_schedule["tuesday"]
    elif today.weekday() == 2:
        schedule_info = class_schedule["wednesday"]
    elif today.weekday() == 3:
        schedule_info = class_schedule["thursday"]
    elif today.weekday() == 4:
        schedule_info = class_schedule["friday"]
    else:
        schedule_info = ["No classes today!"]

    if schedule_info:
        # ANSI escape codes for text formatting
        start_color = "\033[1;33m"  # Yellow color for highlighting
        end_color = "\033[0m"  # Reset color

        print(f"{start_color}\033[4mSchedule for Today:{end_color}")
        for item in schedule_info:
            print(f"{start_color}{item}{end_color}")


def get_user_input():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source)  # Adjust for ambient noise
        audio = r.listen(source)

    try:
        recognized_text = r.recognize_google(audio).lower()
        print("Recognized Text:", recognized_text)  # Debug print
        return recognized_text
    except sr.UnknownValueError:
        print("Sorry, I could not understand audio. Please try again.")
        return get_user_input()
    except sr.RequestError as e:
        print(f"Error connecting to Google API: {e}")
        return ""


def open_word_document(file_name):
    full_path = r"C:\Users\wilso\OneDrive"
    try:
        subprocess.Popen(["start", "winword", f"{full_path}\\{file_name}"], shell=True)
    except Exception as e:
        print(f"Error opening Word document: {e}")


def extract_file_name(tokens):
    start_index = None
    end_index = None

    for i in range(len(tokens)):
        if tokens[i].lower() == "open" and i < len(tokens) - 1:
            if tokens[i + 1].lower() == "the":
                start_index = i + 2
            else:
                start_index = i + 1
        elif tokens[i].lower() == "word":
            end_index = i
            break

    if start_index is not None and end_index is not None:
        file_name = " ".join(tokens[start_index:end_index])
        return file_name.lower()
    else:
        return ""

def extract_song_name(tokens):
    start_index = None
    end_index = None

    for i in range(len(tokens)):
        if tokens[i].lower() == "play" and i < len(tokens) - 1:
             start_index = i + 1
        elif tokens[i].lower() == "by":
             end_index = i
             break

    if start_index is not None and end_index is not None:
       song_name = " ".join(tokens[start_index:end_index])
       return song_name.lower()
    else:
       return ""


def extract_artist_name(tokens):
    start_index = None
    end_index = len(tokens)

    for i in range(len(tokens)):
        if tokens[i].lower() == "by" and i < len(tokens) - 1:
             start_index = i + 1

    if start_index is not None:
       song_name = " ".join(tokens[start_index:end_index])
       return song_name.lower()
    else:
       return ""


def fetch_recent_emails(username, password):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using your credentials
        my_mail.login(username, password)

        # Select the Inbox to fetch messages
        my_mail.select('Inbox')

        # Search for all emails and fetch the 10 most recent ones, sorted by date
        _, data = my_mail.search(None, 'ALL', 'SINCE',
                                 datetime.date.today().strftime('%d-%b-%Y'))  # Search for emails from today
        mail_id_list = data[0].split()  # IDs of all emails

        # Take only the first 10 emails

        emails_info = []

        for num in mail_id_list:
            _, msg_data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            subject = msg['subject']
            sender = msg['from']
            emails_info.append({'subject': subject, 'sender': sender})

        my_mail.close()
        my_mail.logout()

        return emails_info
    except Exception as e:
        print(f"Error fetching emails: {e}")
        return []


def check_email():


    # Call the method to fetch recent emails and print them
    my_email = "wilsondu88@gmail.com"
    my_password = ""
    recent_emails = fetch_recent_emails(my_email, my_password)

    if not recent_emails:
        print("No recent emails found.")
        return

    print("Recent Emails:")
    for i, email_info in enumerate(recent_emails, start=1):
        print(f"Email {i}:")
        print(f"Subject: {email_info['subject']}")
        print(f"Sender: {email_info['sender']}")
        print("_________________________________________")


def find_closest_match(query, choices):
    # Use fuzzywuzzy to find the closest match
    result, score = process.extractOne(query, choices)

    threshold = 70  # Adjust as needed

    if score >= threshold:
        return result
    else:
        return None


def play_spotify_track(track_name, artist_name):
    # Set up Spotify API credentials
    sp_oauth = SpotifyOAuth(client_id='e98baec2f1e64579aae9aba92f55eb83',
                            client_secret='81eaa189587e44caa1e25aac898f8b76',
                            redirect_uri='https://google.com',
                            scope='user-library-read user-modify-playback-state user-read-playback-state user-modify-playback-state')

    # Open the authorization URL in the default web browser
    auth_url = sp_oauth.get_authorize_url()
    webbrowser.open(auth_url)

    # Prompt user to paste the redirected URL
    redirected_url = input("Please paste the redirected URL here: ")

    # Get access token from the redirected URL
    token_info = sp_oauth.get_access_token(redirected_url)

    # Set up Spotify API with obtained access token
    sp = spotipy.Spotify(auth=token_info['access_token'])

    # Search for the track
    playlist_id = '5DBVdxF3R16Twha4AszHSX'
    playlist_info = sp.playlist_tracks(playlist_id)

    # Extract track names and artist names from the playlist
    available_track_names = [track['track']['name'] for track in playlist_info['items']]
    available_artist_names = [artist['name'] for track in playlist_info['items'] for artist in track['track']['artists']]
    results = sp.search(q=f'{track_name} {artist_name}', type='track', limit=1)
    fuzzy_track_name = find_closest_match(track_name, available_track_names)
    fuzzy_artist_name = find_closest_match(artist_name, available_artist_names)

    # Check if any tracks were found
    if fuzzy_track_name and fuzzy_artist_name:
        # Search for the track using fuzzy-matched values
        results = sp.search(q=f'{fuzzy_artist_name} {fuzzy_track_name}', type='track', limit=1)

        # Check if any tracks were found
        if results['tracks']['items']:
            track_uri = results['tracks']['items'][0]['uri']

            # Start playback

            sp.start_playback(uris=[track_uri])

            print(f"Playing {fuzzy_track_name} by {fuzzy_artist_name} on Spotify.")
        else:
            print(f"Could not find a close match for the track {track_name} by {artist_name} on Spotify.")
    else:
        print(f"Could not find a close match for the track {track_name} by {artist_name} on Spotify.")


def process_user_input(user_input):
    tokens = word_tokenize(user_input)
    tagged_tokens = pos_tag(tokens)

    print("Tokenized Input:")
    for token, pos in tagged_tokens:
        print(f"{token}: {pos}")

    # Modify keyword checks to be case-insensitive
    if "email" in [t.lower() for t in tokens] or "gmail" in [t.lower() for t in tokens]:
        if "open" in [t.lower() for t in tokens]:
            open_gmail()
        else:
            check_email()
    elif "youtube" in [t.lower() for t in tokens]:
        open_youtube()
    elif "outlook" in [t.lower() for t in tokens]:
        open_outlook()
    elif "schedule" in [t.lower() for t in tokens]:
        schedule()
    elif "canvas" in [t.lower() for t in tokens]:
        open_canvas()
    elif "discord" in [t.lower() for t in tokens]:
        if "close" in [t.lower() for t in tokens]:
            close_discord()
        else:
            open_discord()
    elif "kindle" in [t.lower() for t in tokens]:
        if "close" in [t.lower() for t in tokens]:
            close_kindle()
        else:
            open_kindle()
    elif "spotify" in [t.lower() for t in tokens]:
        if "close" in [t.lower() for t in tokens]:
            close_spotify()
        else:
            open_spotify()
    elif "word" in [t.lower() for t in tokens]:
        file_name = ""
        file_name = extract_file_name(tokens)
        if file_name != "":
            # Use difflib to get close matches
            close_matches = difflib.get_close_matches(file_name, existing_word_docs, n = 1, cutoff=.3)

            if close_matches:
                suggested_file_name = close_matches[0] + ".docx"
                open_word_document(suggested_file_name)

            else:
                print("No close matches found. Please specify the Word document you want to open.")
        else:
            print("Please specify the Word document you want to open.")
    elif "play" in [t.lower() for t in tokens]:
        song_name = ""
        song_name = extract_song_name(tokens)
        artist_name = ""
        artist_name = extract_artist_name(tokens)
        if artist_name != "" and song_name != "":
            play_spotify_track(song_name, artist_name)
    else:
        print("I'm sorry, I cannot perform this function.")

if __name__ == '__main__':
    print("Hello Brian, what can I assist you with right now?")

    while True:
        user_input = get_user_input()

        if user_input == "quit":
            break

        process_user_input(user_input)

        print("Is there anything else I can do for you?")


