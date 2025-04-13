import win32com.client
import speech_recognition as sr
import webbrowser
from datetime import datetime

# Initialize the speaker
speaker = win32com.client.Dispatch("SAPI.SpVoice")




# Function to recognize voice input
def takeCommand():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        audio = recognizer.listen(source)

        try:
            print("Recognizing...")
            query = recognizer.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, I did not catch that.")
            return ""
        except sr.RequestError:
            print("Internet connection error.")
            return "internet issue"


# Function to speak output
def speak(text):
    speaker.Speak(text)

def play_song(song_name):
    speak(f"Playing {song_name} on YouTube...")
    webbrowser.open(f"https://www.youtube.com/results?search_query={song_name}")


# Main loop
def NEXUS():
    speak("Hello, I am NEXUS AI. How can I help you?")
    
    # Website dictionary
    sites = {
        "google": "https://www.google.com",
        "youtube": "https://www.youtube.com",
        "facebook": "https://www.facebook.com",
        "twitter": "https://www.twitter.com",
        "github": "https://www.github.com",
        "linkedin": "https://www.linkedin.com",
        "instagram": "https://www.instagram.com"
    }

    while True:
        command = takeCommand()

        # Exit condition
        if "exit" in command or "stop" in command or "quit" in command:
            speak("Goodbye! Have a nice day.")
            print("Exiting...")
            break

        # Opening websites
        site_opened = False
        for site in sites:
            if site in command:
                speak(f"Opening {site}...")
                webbrowser.open(sites[site])
                site_opened = True
                break  # Open only once, then exit the loop

        if site_opened:
            continue  # Go back to listening for the next command

        # Respond to common queries
        if "how are you" in command:
            speak("I am doing great! How about you?")

        elif "time" in command:
            current_time = datetime.now().strftime("%H:%M:%S")
            speak(f"The current time is {current_time}")

        elif "play" in command:
            song = command.replace("play", "").strip()
            play_song(song)

        elif command:
            speak(f"You said: {command}")


# Run the assistant
NEXUS()
