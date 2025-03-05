import os
import subprocess
import sys
import ctypes
import speech_recognition as sr
import pyttsx3

base_directory = r''  # Add your directory path, use / to separate locations

def speak(text):
    """Convert text to speech."""
    engine = pyttsx3.init("sapi5")
    voices = engine.getProperty("voices")
    engine.setProperty("voice", voices[1].id)
    engine.setProperty("rate",180)
    engine.say(text)
    engine.runAndWait()

def listen():
    """Capture verbal input from the user."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
        try:
            return recognizer.recognize_google(audio).lower()
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            return ""
        except sr.RequestError:
            print("Could not request results from Google Speech Recognition service.")
            return ""

def list_files_in_directory(directory):
    """List all files in the given directory."""
    try:
        return os.listdir(directory)
    except FileNotFoundError:
        print(f"Directory {directory} not found.")
        return []

def choose_folder():
    """Prompt the user to choose between 'enable' and 'disable' folders verbally."""
    speak("Which folder do you want to open? Please say enable or disable.")
    print("Available folders: enable, disable")
    while True:
        choice = listen()
        if choice in ['enable', 'disable']:
            return choice
        else:
            speak("Invalid choice. Please say enable or disable.")

def choose_registry_key(files):
    """Prompt the user to choose a registry key from the list of files verbally."""
    print("\nAvailable registry keys:")
    for index, file in enumerate(files):
        print(f"{index + 1}: {file}")
    
    speak("Choose a registry key by number or say the key name.")
    
    while True:
        choice = listen()
        
        if choice.isdigit() and 0 < int(choice) <= len(files):
            return files[int(choice) - 1]
        elif choice in files:
            return choice
        else:
            speak("Invalid choice. Please try again.")

def run_registry_key(file_path):
    """Run the selected registry key file."""
    if not file_path.endswith('.reg'):
        print("Selected file is not a valid .reg file.")
        return
    
    # This finds the registry key in the directory
    registry_key_name = os.path.basename(file_path)
    

    speak(f"Preparing to import the registry key: {registry_key_name}")
    
    try:
        subprocess.run(['regedit', '/s', file_path], check=True)
        speak(f"Successfully imported: {registry_key_name}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to import registry key: {e}")

def is_admin():
    """Check if the script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    """Main function to run the module."""
    folder_choice = choose_folder()
    folder_path = os.path.join(base_directory, folder_choice)
    
    files = list_files_in_directory(folder_path)
    
    if files:
        selected_key = choose_registry_key(files)
        selected_key_path = os.path.join(folder_path, selected_key)
        print(f"You selected: {selected_key}")
        
        # This selects the key to run it
        run_registry_key(selected_key_path)
    else:
        speak(f"No files found in the {folder_choice} folder.")

if __name__ == "__main__":
    if not is_admin():
        print("This script requires administrative privileges. Please allow the UAC prompt.")
        subprocess.run(['powershell', '-Command', 'Start-Process', sys.executable, '-ArgumentList', '"{}"'.format(' '.join(sys.argv)), '-Verb', 'runAs'])
        sys.exit()  
    else:
        main()
