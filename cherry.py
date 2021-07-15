import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import smtplib
import subprocess
from progress.bar import ChargingBar
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import ctypes
import pyjokes

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[1].id)  #put 1 for female voice
engine.setProperty('voice', voices[1].id) #put 1 for female voice


def speak(audio):           #it will speak whatever we type
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning !")

    elif hour>=12 and hour<18:
        speak("Good Afternoon !")

    else:
        speak("Good Evening !")

    speak("I am cherry. Your Personal Assistant, Please tell me What can i do for you.")

def takeCommand():  #it takes microphone input from the user and returns string o/p
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listenning ...")
        r.pause_threshold = 1    #pause the audio for seconds
        audio = r.listen(source)

    try:
        print("Recognizing ...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        #print(e)
        print("Didn't Recognize, Please say again...")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremailid@gmail.com', 'password')
    server.sendmail('youremailid@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    clear = lambda: os.system('cls')
    clear()               #this function clears all commands before executing the program
    wishMe()
    while True:
    #if 1:        #it will listen one time only 
        query = takeCommand().lower()

        #Below is the logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open chrome' in query:
            speak('Opening Chrome...')
            chromePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(chromePath)

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or "play songs" in query:
            music_dir = 'C:\\Users\\AATIQ\\Music'
            songs = os.listdir(music_dir)
            print(songs)
            n = random.randint(0, 10)
            print(n)
            os.startfile(os.path.join(music_dir, songs[n]))

        elif 'the time' in query: 
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")

        elif 'open Visual Studio Code' in query:
            codePath = "C:\\Users\\AATIQ\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif "who are you" in query:
            speak("I am your virtual assistant Sir")

        elif "who  made you" in query or "who created you" in query:
            speak("I have been created by Aatiq.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'email to a' in query:
            try:
                speak("What should i say?")
                content = takeCommand()
                to = "atq.rock@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("Failed to send mail, Please try again...")
                #to send email u nees to turn on "less secure app access" on google

        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
                    
        elif 'exit' in query or 'quit' in query or 'stop' in query:
            speak("Thanks for giving me your time")
            exit()