import speech_recognition as sr
import datetime,time
import pyttsx3
import webbrowser
import os
import subprocess
import requests
from googlesearch import search
from bs4 import BeautifulSoup
import pythoncom
import sys
import re

def get_d_drive_directories():
    d_drive_dirs = []
    try:
        for item in os.listdir("D:/"):
            full_path = os.path.join("D:/", item)
            if os.path.isdir(full_path):
                d_drive_dirs.append(full_path)
    except FileNotFoundError:
        print("D: drive not found")
    return d_drive_dirs

SEARCH_DIRECTORIES = [
    "D:/",
    "C:/",
    "C:/Users/" + os.getlogin(),
    "C:/Users/" + os.getlogin() + "/Documents",
    "C:/Users/" + os.getlogin() + "/Desktop",
    "C:/Users/" + os.getlogin() + "/Downloads",
]

SEARCH_DIRECTORIES.extend(get_d_drive_directories())

def speak(text):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.setProperty('rate', 150)
    engine.say(text)
    engine.runAndWait()

def listen(energy_threshold=600, pause_threshold=0.8, phrase_time_limit=20):
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = energy_threshold
    recognizer.pause_threshold = pause_threshold

    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        try:
            audio_data = recognizer.listen(source, timeout=3, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            print("No speech detected")
            return ""
    
    try:
        print("Recognizing...")
        text = recognizer.recognize_google(audio_data, language="en-US")
        print(f"You said: {text}")
        return text.lower()
    except sr.UnknownValueError:
        print("Speech recognition could not understand audio")
        return ""
    except sr.RequestError as e:
        print(f"Error occurred while requesting the recognition service: {e}")
        return ""

def get_folder_name():
    for _ in range(3):
        speak("Please say the name of the folder you'd like to open.")
        folder_name = listen(energy_threshold=600, phrase_time_limit=15)
        if folder_name:
            return folder_name
        else:
            speak("I'm sorry, I didn't catch that. Let's try again.")
    
    speak("I'm having trouble understanding the folder name. Please type the folder name.")
    return input("Please type the folder name: ")

def wish():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        greeting = "Good Morning Sir"
    elif 12 <= hour < 17:
        greeting = "Good Afternoon Sir"
    else:
        greeting = "Good Evening Sir"
    speak(greeting)
    speak("I'm Jarvis, your virtual assistant. Just say something when you need me.")
        

def open_drive(drive_letter):
    drive_path = f"{drive_letter}:\\"
    if os.path.exists(drive_path):
        subprocess.Popen(f'explorer "{drive_path}"')
        speak(f"Opening {drive_letter} drive")
    else:
        speak(f"Sorry, {drive_letter} drive is not accessible")    


def clean_filename(filename):
    # Replace common speech-to-text artifacts
    filename = filename.lower()
    filename = filename.replace(" dot ", ".")
    filename = filename.replace("dot ", ".")
    filename = filename.replace(" period ", ".")
    filename = filename.replace("period ", ".")
    
    # Remove any spaces before file extensions
    filename = re.sub(r'\s+([.]\w+)$', r'\1', filename)
    
    # Remove any non-filename characters
    filename = re.sub(r'[^\w.-]', '', filename)
    
    return filename

def find_file(filename):
    cleaned_filename = clean_filename(filename)
    for drive in ['C:', 'D:']:  # Add more drives if needed
        for root, dirs, files in os.walk(drive + '\\'):
            for file in files:
                if cleaned_filename.lower() in file.lower():
                    return os.path.join(root, file)
    return None

def find_folder(foldername):
    for directory in SEARCH_DIRECTORIES:
        for root, dirs, files in os.walk(directory):
            if foldername.lower() in (dir.lower() for dir in dirs):
                return os.path.join(root, foldername)
            for dir in dirs:
                if foldername.lower() in dir.lower():
                    return os.path.join(root, dir)
    return None

def get_google_search_results(query):
    search_results = search(query, num_results=5)
    return list(search_results)

def extract_info_from_url(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        paragraphs = soup.find_all('p')
        text = ' '.join([para.get_text() for para in paragraphs])
        return text
    except Exception as e:
        print(f"Error extracting information from URL {url}: {e}")
        return ""

def get_relevant_info(text):
    sentences = text.split('. ')
    if len(sentences) > 0:
        return sentences[0]
    return "Sorry, I couldn't find specific information."

def open_url(url, site_name):
    print(f"Attempting to open {site_name}...")
    speak(f"Opening {site_name}")
    
    try:
        webbrowser.open(url, new=2)
        print(f"Opened {site_name} using webbrowser")
    except Exception as e:
        print(f"Error with webbrowser: {e}")
        try:
            os.system(f"start {url}")
            print(f"Opened {site_name} using os.system")
        except Exception as e:
            print(f"Error with os.system: {e}")
            try:
                subprocess.Popen(['start', url], shell=True)
                print(f"Opened {site_name} using subprocess")
            except Exception as e:
                print(f"Error with subprocess: {e}")
                speak(f"I'm having trouble opening {site_name}. Please try manually.")
                return False
    return True

def virtual_assistant():
    wish()
    
    while True:
        command = listen()
        if command:
            if "open c drive" in command:
                open_drive("C")
            elif "open d drive" in command:
                open_drive("D")
            elif "open google" in command:
                open_url("https://www.google.com", "Google")
            elif "open youtube" in command:
                open_url("https://www.youtube.com", "YouTube")
            elif "open whatsapp" in command:
                open_url("https://web.whatsapp.com", "WhatsApp")
            elif "open instagram" in command:
                open_url("https://www.instagram.com", "Instagram")
            elif "open github" in command:
                open_url("https://github.com/CHIRAGCHHONKAR", "GitHub")
            elif "open claude" in command or "open cloud ai" in command or "claude ai" in command:
                open_url("https://claude.ai", "Claude AI")
            elif "open chat gpt" in command:
                open_url("https://chat.openai.com/", "ChatGPT")
            elif "open freepik" in command:
                open_url("https://www.freepik.com/", "Freepik")
            elif "open supabase" in command:
                open_url("https://supabase.com/dashboard/projects", "Supabase")
            elif "open vs code" in command:
                speak("Opening VS Code")
                os.system("code")  
            elif "what is your name" in command:
                speak("My name is Jarvis, your virtual assistant.")
            elif "what time is it" in command:
                current_time = datetime.datetime.now().strftime("%H:%M")
                speak(f"The current time is {current_time}")
            elif "what is today's date" in command:
                current_date = datetime.datetime.now().strftime("%Y-%m-%d")
                speak(f"Today's date is {current_date}")
            elif "shutdown" in command:
                speak("Shutting down the computer.")
                os.system("shutdown /s /t 1")
            elif "restart" in command:
                speak("Restarting the computer.")
                os.system("shutdown /r /t 1")
            elif "list files" in command:
                speak("Which directory would you like to list files from?")
                directory = listen()
                if directory:
                    try:
                        files = os.listdir(os.path.expanduser(directory))
                        speak(f"Listing files in {directory}")
                        for file in files:
                            print(file)
                            speak(file)
                    except:
                        speak(f"Sorry, I couldn't access the directory {directory}")
            elif "open file" in command:
                speak("What's the name of the file you'd like to open?")
                filename = listen()
                if filename:
                    cleaned_filename = clean_filename(filename)
                    speak(f"Searching for file: {cleaned_filename}")
                    file_path = find_file(cleaned_filename)
                    if file_path:
                        try:
                            os.startfile(file_path)
                            speak(f"Opening {(filename)}")
                        except Exception as e:
                            speak(f"Sorry, I couldn't open {cleaned_filename}. Error: {str(e)}")
                    else:
                        speak(f"Sorry, I couldn't find {cleaned_filename}")
                else:
                    speak("I'm sorry, I didn't catch the filename. Please try again.")
            elif "open folder" in command:
                folder_name = get_folder_name()
                if folder_name:
                    speak(f"Searching for folder: {folder_name}")
                    folder_path = find_folder(folder_name)
                    if folder_path:
                        try:
                            if sys.platform == "win32":
                                os.startfile(folder_path)
                            else:
                                subprocess.Popen(["xdg-open", folder_path])
                            speak(f"Opening folder {os.path.basename(folder_path)}")
                        except Exception as e:
                            speak(f"Sorry, I couldn't open the folder {folder_name}. Error: {str(e)}")
                    else:
                        speak(f"Sorry, I couldn't find the folder {folder_name}")
                else:
                    speak("I'm sorry, I couldn't get the folder name. Please try again later.")
            elif "who is" in command or "what is" in command:
                search_results = get_google_search_results(command)
                if search_results:
                    first_result = search_results[0]
                    info = extract_info_from_url(first_result)
                    relevant_info = get_relevant_info(info)
                    speak(relevant_info)
                else:
                    speak("Sorry, I couldn't find information for that query.")
            elif "exit" in command:
                speak("Goodbye, Sir!")
                sys.exit()
        else:
            time.sleep(0.1)

def create_startup_batch():
    batch_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run_jarvis.bat")
    with open(batch_path, "w") as batch_file:
        batch_file.write(f'@echo off\n')
        batch_file.write(f'start "" pythonw "{os.path.abspath(__file__)}"\n')
    print(f"Batch file created at: {batch_path}")
    return batch_path

if __name__ == "__main__":
    create_startup_batch()  # This will create/update the batch file every time
    pythoncom.CoInitialize()
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.setProperty('rate', 150)
    virtual_assistant()  