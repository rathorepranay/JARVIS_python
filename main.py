from http import client
import os
import random
from win32com.client import Dispatch
import speech_recognition as sr
import webbrowser
import datetime
import openai

# Initialize text-to-speech
speaker = Dispatch("SAPI.SpVoice")



def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = .5
        try:
            audio = r.listen(source)
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            say("Sorry, I could not understand the audio.")
            return ""
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            return ""


if __name__ == "__main__":
    say("Hello Sir!")
    while True:
        command = takeCommand()
        if command[0] == open:
            second = command[1]
            webbrowser.open(f"https://{second}.com/")

        elif "stop listening" in command.lower():
            say("Stopped listening. Goodbye!")
            break
        elif "the time" in command.lower():
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir, the time is {current_time}")
        elif "chat gpt" in command.lower():
            response = ai(prompt=command)
            say("Here is the response from OpenAI.")
            print(response)
