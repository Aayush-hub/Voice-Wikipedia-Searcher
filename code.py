import pyttsx3                    #pip install pyttsx3 - module to convert text to speech
import datetime                   #built-in module
import speech_recognition as sr   #pip install speechRecognition
import webbrowser                 #built-in module
import wikipedia                  #pip install wikipedia
import os
engine = pyttsx3.init('sapi5')    #microsoft sound API
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  #setting Zira's voice


def speak(str):                                              #converting text to speech....
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)


def wish():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir! How may I help you!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir! How may I help you!")

    else:
        speak("Good Evening Sir! How may I help you!")

def takeCommand():                                            #converting users input to text and searching by google API

    r=sr.Recognizer()
    # print(sr.Microphone.list_microphone_names())
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source,duration=1)
        # r.energy_threshold(300)
        print("say anything : ")
        audio= r.listen(source)
    try:
        query = r.recognize_google(audio)
        print(query)
    except:
        print("sorry, could not recognise")
        return "None"
    return query


if __name__ == "__main__":
    wish()
    
    while True:
        query = takeCommand().lower()         #converting speech (userinput) to text and searching on wikipedia
        if 'wikipedia' in query:
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=4)
            print(results)
            speak(f"According to Wikipedia: {results}")
        elif 'open google' in query:      #opening wikipedia.com
            webbrowser.open("wikipedia.com")
