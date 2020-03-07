import re  # pip install re
import urllib  # pip install urllib
import win32com.client as wincl  # pip install win32
import speech_recognition as sr  # pip install speechRecognition
import datetime
import wikipedia  # pip install wikipedia
import webbrowser
import os
import smtplib  # pip install smtplib
import random
import sys
import requests

s = wincl.Dispatch('SAPI.SpVoice')
s.Voice = s.GetVoices().Item(1)
s.Speak("Hii boss")

info = '''
       +=======================================+
       |.......LYRA VIRTUAL INTELLIGENCE.......|
       +---------------------------------------+
       |#Author: abhishek raut                 |
       |#Date: 01/02/2020                      |
       |#sub: python programming               |
       +---------------------------------------+
       |......LYRA VIRTUAL INTELLIGENCE........|
       +=======================================+
       |              OPTIONS:                 |
       |#hello/hi     #goodbye    #sleep mode  |
       |#your name    #LYRA       #what time   |
       |#open google  #sing song  #music       |
       |#search youtube  #search wikipedia     |
       |#start/stop someapp  #ask for jokes    |
       |#search google #etc                    |
       +=======================================+
       
       '''


def speak(audio):
    s.say(audio)
    s.runAndWait()


def wish():  # wishes according to time
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        s.speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        s.speak("Good Afternoon!")

    else:
        s.speak("Good Evening!")

    s.speak("I am Lyra Sir. Please tell me how may I help you")
    print(info)


def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said:,{query}\n")
        return query

    except Exception as e:
        # print(e)
        print("Say that again please...\n")
        return "none"


def your_results():  # returns voice messege after command successfully worked
    s.speak("here's the your results sir")


if __name__ == "__main__":
    wish()  # wish function call
    while True:
        query = takeCommand().lower()  # converts input into lowercase and store into query variable
        # Logic for executing tasks based on query

        if ' search wikipedia' in query:
            s.speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            s.speak("According to Wikipedia")
            print(results)
            s.speak(results)

        elif "ok lyra" in query or "hey lyra" in query:
            okay = ['Yes Sir?', 'What can I doo for you sir?']
            s.speak(random.choice(okay))

        elif 'hello' in query or 'hi' in query:
            ok = ['Yes Sir?', 'What can I doo for you sir?', 'Hello! I am Lyra. Give me a command']
            s.speak(random.choice(ok))

        elif "youtube and search" in query:
            s.speak('Ok!')
            reg_ex = re.search('youtube (.+)', query)
            if reg_ex:
                domain = query.split("youtube", 1)[1]
                query_string = urllib.parse.urlencode({"search_query": domain})
                html_content = urllib.request.urlopen("http://www.youtube.com/results?" + query_string)
                search_results = re.findall(r'href=\"\/watch\?v=(.{11})',
                                            html_content.read().decode())  # finds all links in search result
                webbrowser.open("http://www.youtube.com/watch?v={}".format(search_results[0]))
                your_results()

        elif 'open youtube' in query or 'goto youtube' in query:
            webbrowser.open("youtube.com")
            your_results()

        elif 'open wikipedia' in query :
            webbrowser.open("wikipedia.com" + query)
            your_results()

        elif "today's weather" in query or "current weather" in query or "weather" in query:
            webbrowser.open("https://www.accuweather.com/en/in/nagpur/204844/weather-forecast/204844")
            your_results()

        elif 'open facebook' in query or 'open website facebook' in query or 'open facebook website' in query or 'goto facebook' in query:
            webbrowser.open("facebook.com")
            your_results()

        elif 'start google' in query:
            webbrowser.open("google.com")
            your_results()

        elif ' open google and search' in query or 'search on google' in query:
            url = "https://www.google.co.in/search?q=" + query
            webbrowser.open_new(url)
            your_results()

        elif "open map" in query:
            webbrowser.open("https://www.google.com/maps")
            your_results()

        elif "open gmail" in query:
            webbrowser.open(
                "https://accounts.google.com/signin/v2/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin")
            your_results()

        elif 'open stackoverflow' in query or "open stack over flow" in query or "open stack overflow" in query or "open stackover flow" in query:
            webbrowser.open("stackoverflow.com")
            your_results()

        elif 'play music' in query or 'play song' in query:
            music_dir = 'F:\\abhishek\\songs'  # change path where you've save musics
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))
            s.speak("here's the your results sir, enjoy your favroute song")

        elif "open chrome" in query:
            s.speak("Google Chrome")
            os.startfile(
                'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')  # change the code path to access software
            your_results()

        elif "firefox" in query or "mozilla" in query:
            s.speak("Opening Mozilla Firefox")
            os.startfile(
                'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')  # change the code path to access software
            your_results()

        elif "word" in query:
            s.speak("Opening Microsoft Word")
            os.startfile(
                'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\ Programs\\Microsoft Office 2013\\Word 2013.lnk')  # change the code path to access software
            your_results()

        elif 'the time' in query or 'time is' in query or 'a time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            s.speak(f"Sir, the time is {strTime}")

        elif 'open pcsx2' in query:
            s.speak("Opening pcsx2")
            codePath = 'C:\\Program Files (x86)\\PCSX2 1.4.0\\pcsx2.exe'  # change the code path to access software
            os.startfile(codePath)
            your_results()

        elif 'open pycharm' in query:
            s.speak("Opening pycharm application")
            codePath = 'C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.3.3\\bin\\pycharm64' + ".exe  "  # change the code path to access software
            os.startfile(codePath)
            your_results()

        elif 'open calculator' in query:
            s.speak("Opening Calculator")
            codePath = 'C:\\Windows\\System32\\calc.exe'  # change the code path to access software
            os.startfile(codePath)
            your_results()

        elif "who made you" in query or "created you" in query:
            s.speak("I have been created by my master abhishek")

        elif "your name" in query:
            s.speak("my name is lyra sir. how did you forgot?")

        elif "how are you" in query or "are you ok" in query:
            s.speak("im fine. how about you sir!!")

        elif "who are you" in query or "define yourself" in query:
            sp = '''Hello. Im lyra your personal Assistant. 
                   I am here to make your life easier. 
                   You can command me to perform various tasks such as searching wikipedia or opening applications etcetra'''
            s.speak(sp)
        elif "where are you from" in query:
            s.speak("basically im from my master abhishek computer")

        elif "what are you doing" in query:
            s.speak("nothing right now!. but i can make your life easy , haha")

        elif "how can you help me" in query:
            s.speak("i can help you to open apps , play music, search on web !!!,"
                    "just tell me to do something")

        elif "why are you created" in query:
            s.speak("i have been created to make you happy. hahaha!!")

        elif "what can you do for me" in query:
            ls = ['i can sing song for you. if you dont mind!!',
                  'i can tell you jokes!!. if you dont mind!!',
                  'i can search youtube for you!!. if you dont mind!!',
                  'i can play music for you!!. if you dont mind!!']

            s.speak(random.choice(ls))

        elif "sing a song" in query or "sing song" in query or "sing a birthday song" in query:
            s.speak(
                "Happy Birthday to You , cha, cha, cha, Happy Birthday to You, cha, cha, cha,  Happy Birthday Dear abhishek "
                " Happy Birthday to You, cha, cha, cha.")

        elif "geeksforgeeks" in query:
            s.speak("Geeks for Geeks is the Best Online Coding Platform for learning.")

        elif "jokes" in query or "joke" in query:
            jokes = ['Can a kangaroo jump higher than a house? Of course, a house doesn’t jump at all.',
                     'My dog used to chase people on a bike a lot. It got so bad, finally I had to take his bike away.',
                     'Doctor: Im sorry but you suffer from a terminal illness and have only 10 to live.Patient: What do you mean, 10? 10 what? Months? Weeks?!"Doctor: Nine.',
                     ' Okay, here you go. What do you call a guy with a rubber toe? Roberto.']
            s.speak(random.choice(jokes))

        elif "exit" in query or "bye" in query or "sleep" in query or "shutdown" in query:
            s.speak("Ok bye sir, have a nice day")
            sys.exit()

        elif "power off" in query or "power of" in query or "turn of" in query or "turn off" in query:
            s.speak('good bye. lyra powering off in 4, 3, 2, 1, 0')
            sys.exit()

        elif "thanks" in query or "thank you" in query:
            s.speak("your welcome")

        elif "show me info" in query or "give me info" in query:
            print(info)
            your_results()

        else:
            ls = ["Try to ask google", "Provide correct command", "Search could not found", "Try to ask correctly",
                  "You are awesome but sry i cannot find results"]
            print(random.choice(ls))

