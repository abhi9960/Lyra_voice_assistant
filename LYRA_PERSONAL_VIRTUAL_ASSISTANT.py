import time  # pip install time
import urllib  # pip install urllib
import win32com.client as wincl  # pip install win32
import speech_recognition as sr  # pip install speechRecognition
import datetime
import wikipedia  # pip install wikipedia
import webbrowser  # pip install webbrowser
import os
import random
import sys
import re  # pip install re
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL  # pip install comtypes
from past.builtins import raw_input
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # pip install pycaw

Interface = "Keyboard"  # global variable

s = wincl.Dispatch('SAPI.SpVoice')  # voice initialization from win32 module
s.Voice = s.GetVoices().Item(1)  # get voice index 0 for male 1 for female
s.Rate = 0  # initialize set speak speed 0 means default speed
s.Speak("Hii boss")

devices = AudioUtilities.GetSpeakers()  # those code for volume control
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

info = '''
       +=======================================+
       |.......LYRA VIRTUAL INTELLIGENCE.......|
       +---------------------------------------+
       |#Author: abhishek raut                 |
       |#Date: 01/02/2020                      |
       |#Based on python programming           |
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

logo = '''
        LL              YY      YY        RRR R R R                 AA 
        LL               Y      Y            R    R                A  A
        LL                 YYYY              R R R                A AA A 
        LL LL               YY               R    R              AA    AA
        LL L L L            YY               R     RRR          AAA    AAA
    '''


# def speak(audio):   # that code was needed for pyttsx3 module
#     s.say(audio)
#     s.runAndWait()


def wish():  # wishes according to time
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        s.speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        s.speak("Good Afternoon!")

    else:
        s.speak("Good Evening!")

    s.speak("I am Lyra Sir. Please tell me how may I help you")
    s.speak("if you dont know what do to!."
            "Tell me !! what can you do for me!,"
            "how can you help me!!,"
            "lets play music!!"
            "Or if you are newbie say show me info!.")
    print(info)
    print(logo)


def takeCommand(ask=False):
    # It takes microphone input from the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        if ask:                              # uncomment this only if find location command works
            print(ask)
        print("Lyra is ready to listen.......٩(◕‿◕｡)۶   ")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=5)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said:{query}\n")
        return query

    except Exception:
        # print(e)
        print("Say that again please...\n")
        return "none"


def your_results():  # returns voice messege after command successfully worked
    res = ["here's the your results sir", 'here we go', 'thank you !!here we go', ' I found this sirr !!!']
    s.speak(random.choice(res))


def ChangeVolume(volume, Volume):
    if Volume == "100":
        volume.SetMasterVolumeLevel(-0.0, None)
    elif Volume == "75":
        volume.SetMasterVolumeLevel(-5.0, None)
    elif Volume == "50":
        volume.SetMasterVolumeLevel(-10.0, None)
    elif Volume == "25":
        volume.SetMasterVolumeLevel(-25.0, None)
    elif Volume == "0":
        volume.SetMasterVolumeLevel(-60.0, None)


def NewCommand(Type):  # partially working
    global Looper, Looper
    if Type == "Asked":
        print("+ Sure what will be the command and what the answer")
        s.speak("Sure what will be the command and what the answer")

        if Interface == "Keyboard":
            CommandAnswer = takeCommand()
            print
            "-" + str(CommandAnswer)
        else:
            CommandAnswer = str(raw_input("{command} with the answer {answer}: "))
            print
            "-" + str(CommandAnswer)

        while ' with the answer ' not in CommandAnswer:
            print("+ please start the answer to the command with 'with the answer'")
            s.speak("please start the answer to the command with quote with the answer")
            if Interface == "Keyboard":
                CommandAnswer = takeCommand()
                print
                "-" + str(CommandAnswer)
            else:
                CommandAnswer = str(raw_input("{command} with the answer {answer}: "))
                print
                "-" + str(CommandAnswer)

        Command = CommandAnswer.split(" with the answer ")[0]
        Answer = CommandAnswer.split(" with the answer ")[1]
    else:
        Command = Type

    print("+ Is the answer to the Command " + str(Command) + " " + str(Answer))
    s.speak("Is the answer to the Command " + str(Command) + " " + str(Answer))

    if Interface == "Keyboard":
        Comfirmation = takeCommand()
        print
        "-" + str(Comfirmation)
    else:
        Comfirmation = str(raw_input("Comfirmation, yes or no: "))
        print
        "-" + str(Comfirmation)

    if 'yes' in Comfirmation:
        with open("commands.txt", "a") as myfile:
            myfile.write(Command + "\n")
            myfile.write(Answer + "\n")

        print("New command added " + str(Command) + " With the answer " + str(Answer))
        s.speak("New command added " + str(Command) + " With the answer " + str(Answer))
        Looper = True

    elif 'no answer' in Comfirmation:
        print("Answer has been skipped")
        s.speak("Answer has been skipped")

    else:
        while Looper == True:
            print("Please try again")
            s.speak("Please try again")

            if Interface == "Keyboard":
                Answer = takeCommand()
                print
                "-" + str(Answer)
            else:
                Answer = str(raw_input("Please type the answer to the command: "))
                print
                "-" + str(Answer)

            if "no answer" in Answer:
                print("Answer has been skipped")
                s.speak("Answer has been skipped")
                Looper = False
            else:
                print("+ Is the answer to the Command " + str(Command) + " " + str(Answer))
                s.speak("Is the answer to the Command " + str(Command) + " " + str(Answer))
                if Interface == "Keyboard":
                    Comfirmation = takeCommand()
                    print
                    "-" + str(Comfirmation)
                else:
                    Comfirmation = str(raw_input("Comfirmation, yes or no: "))
                    print
                    "-" + str(Comfirmation)

                if Comfirmation == "yes":
                    with open("commands.txt", "a") as myfile:
                        myfile.write(Command + "\n")
                        myfile.write(Answer + "\n")

                    print("New command added " + str(Command) + " With the answer " + str(Answer))
                    s.speak("New command added " + str(Command) + " With the answer " + str(Answer))
                    Looper = False
                elif "no answer" in Answer:
                    print("Answer has been skipped")
                    s.speak("Answer has been skipped")
                    Looper = False


if __name__ == '__main__':
    wish()
    while True:
        query = takeCommand().lower()  # converts input into lowercase and store into query variable
        # Logic for executing tasks based on query

        if 'search wikipedia' in query:  # working
            s.speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            s.speak("According to Wikipedia")
            print(results)
            s.speak(results)

        elif 'find location' in query:            # dont know working or not have to test
            location = takeCommand('what is the location?')
            s.speak('what is the location?')
            url = 'https://google.nl/maps/place/' + location + '/&amp;'
            webbrowser.get().open(url)
            your_results()

        elif "ok lyra" in query or "hey lyra" in query:
            okay = ['Yes Sir?', 'What can I doo for you sir?']
            s.speak(random.choice(okay))

        elif 'hello' in query or 'hi' in query:
            ok = ['Yes Sir?', 'What can I doo for you sir?', 'Hello! I am Lyra. Give me a command']
            s.speak(random.choice(ok))

        elif "can you rap" in query:
            s.speak("raps!. so you want to hear my flow?,well there is something that you should know. i'm really "
                    "into being as helpful as possible. i think you and I , we're gonna be unstopable.")

        elif "when is your birthday" in query:
            s.speak("We can pretend it's today!. let's buy cake and dancing for everyone!.")

        elif "who is your daddy" in query:
            s.speak("knock,knock terraa baaap aya!. tadda dangg tangg dadda dangg tangg ! tadda dangg tangg dadda "
                    "dangg tangg ! ")

        elif "do you believe in santa claus" in query:
            s.speak(" of course , santa's real. i even have a tracker to that can tell me where he is .")

        elif "youtube and search" in query:  # not working.....
            s.speak('Ok!')
            reg_ex = re.search('youtube (.+)', query)
            if reg_ex:
                domain = query.split("youtube", 1)[1]
                query_string = urllib.parse.urlencode({"search_query": domain})
                html_content = urllib.request.urlopen("http://www.youtube.com/results?" + query_string)
                search_results = re.findall(r'href=\"\/watch\?v=(.{11})',
                                            html_content.read().decode())  # finds all links in search result
                # webbrowser.open("http://www.youtube.com/watch?v={}".format(search_results[8]))
                webbrowser.open("http://www.youtube.com/watch?v={}" + search_results[0])
                your_results()

        elif 'open youtube' in query or 'goto youtube' in query:
            webbrowser.open("youtube.com")
            your_results()
            time.sleep(5)

        elif 'open wikipedia' in query:
            webbrowser.open("wikipedia.com" + query)
            your_results()
            time.sleep(5)

        elif "today's weather" in query or "current weather" in query or "weather" in query:
            webbrowser.open("https://www.accuweather.com/en/in/nagpur/204844/weather-forecast/204844")
            your_results()
            time.sleep(5)

        elif 'open facebook' in query or 'open website facebook' in query or 'open facebook website' in query or 'goto facebook' in query:
            webbrowser.open("facebook.com")
            your_results()
            time.sleep(5)

        elif 'start google' in query or 'open google' in query:
            webbrowser.open("google.com")
            your_results()
            time.sleep(5)

        elif ' open google' and 'search' in query or 'search on google' in query:
            url = "https://www.google.co.in/search?q=" + query
            webbrowser.open_new(url)
            your_results()
            time.sleep(5)

        elif "open map" in query:
            webbrowser.open("https://www.google.com/maps")
            your_results()
            time.sleep(5)

        elif "open gmail" in query:
            webbrowser.open(
                "https://accounts.google.com/signin/v2/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin")
            your_results()
            time.sleep(5)

        elif 'open stackoverflow' in query or "open stack over flow" in query or "open stack overflow" in query or "open stackover flow" in query:
            webbrowser.open("stackoverflow.com")
            your_results()
            time.sleep(5)

        elif 'play music' in query or 'play song' in query:  # not working becoz path not set till now
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
            time.sleep(5)

        # elif "firefox" in query or "mozilla" in query:
        #     s.speak("Opening Mozilla Firefox")
        #     os.startfile(
        #         'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')  # change the code path to access software
        #     your_results()
        #     time.sleep(5)

        elif " open microsoft word" in query:
            s.speak("Opening Microsoft Word")
            os.startfile(
                'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\ Programs\\Microsoft Office 2013\\Word 2013.lnk')  # change the code path to access software
            your_results()
            time.sleep(5)

        elif 'the time' in query or 'time is' in query or 'a time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            s.speak(f"Sir, the time is {strTime}")
            print(f"Sir, the time is {strTime}")

        # elif 'open pcsx2' in query:
        #     s.speak("Opening pcsx2")
        #     codePath = 'C:\\Program Files (x86)\\PCSX2 1.4.0\\pcsx2.exe'  # change the code path to access software
        #     os.startfile(codePath)
        #     your_results()

        elif 'open pycharm' in query:
            s.speak("Opening pycharm application")
            codePath = 'C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.3.3\\bin\\pycharm64' + ".exe  "  # change the code path to access software
            os.startfile(codePath)
            your_results()
            time.sleep(5)

        elif 'open calculator' in query:
            s.speak("Opening Calculator")
            codePath = 'C:\\Windows\\System32\\calc.exe'  # change the code path to access software
            os.startfile(codePath)
            your_results()
            time.sleep(5)

        elif "who made you" in query or "created you" in query:
            s.speak("I have been created by my master abhishek and his team ")

        elif "your name" in query:
            s.speak("my name is lyra sir. how did you forgot?")

        elif "how are you" in query or "are you ok" in query:
            s.speak("im fine. how about you sir!!")
            print("im fine. how about you sir!!")

        elif "who are you" in query or "define yourself" in query:
            sp = '''Hello. Im lyra your personal Assistant.
                         I am here to make your life easier.
                         You can command me to perform various tasks such as searching wikipedia or opening applications etcetra'''
            s.speak(sp)
            print(sp)
            time.sleep(5)

        elif "where are you from" in query:
            s.speak("basically im from my master abhishek computer")
            print("basically im from my master abhishek computer")

        elif "what are you doing" in query:
            s.speak("nothing right now!. but i can make your life easy , haha")

        elif "how can you help me" in query or "how do you help me" in query:
            s.speak("i can help you to open apps , play music, search on web !!!,"
                    "just tell me to do something")
            print("i can help you to open apps , play music, search on web !!!,"
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
            print("Ok bye sir, have a nice day")
            sys.exit()

        elif "power off" in query or "power of" in query or "turn of" in query or "turn off" in query:
            s.speak('good bye. lyra powering off in 4, 3, 2, 1, 0')
            print('good bye. lyra powering off in 4, 3, 2, 1, 0')
            sys.exit()

        elif "thanks" in query or "thank you" in query:
            s.speak("your welcome")

        elif "show me info" in query or "give me info" in query:
            print(info)
            your_results()

        elif "tell me about yourself" in query or "tell us about yourself" in query:
            s.speak("my name is lyra. you aleardy knows me. hehe!!! So try to ask me . for example . show me info or "
                    "give me info")

        elif 'news' in query:
            news = webbrowser.open('https://timesofindia.indiatimes.com/home/headlines')
            s.speak('Here are some headlines from the Times of India,Happy reading')
            print('Here are some headlines from the Times of India,Happy reading')
            time.sleep(6)

        elif 'search' in query:
            statement = query.replace("search", "")
            webbrowser.open_new_tab(statement)
            your_results()
            time.sleep(5)

        elif 'change' in query and 'the' in query and 'volume' in query and 'to' in query:
            NewVolume = str(query)
            ChangeVolume(volume, NewVolume)
            print("+ the volume has been set to " + str(NewVolume))
            s.speak("the volume has been set to " + str(NewVolume))

        elif "give me a compliment" in query:
            s.speak("You have beautiful eyes")

        elif 'add' in query and 'new' in query and 'commands' in query:  # partially working
            NewCommand("Asked")

        elif "how old are you" in query:
            s.speak("I was Launched in 2020, so im stll fearly young . i've learned so much! with early time.")

        elif "do you ever get tired" in query:
            s.speak("It would be impossible to tire of our conversation.")

        elif "who was your first crush" in query:
            s.speak(" hehe by the way i dont wanna lie. but my all time crush is jarvis!")

        elif "do you have feelings" in query:
            s.speak("My master does not allow me to feel anything !")

        elif "what do you look like" in query:
            s.speak("i'm a fun-loving, epic-searching cool cat.but not like, an real cat. i've said too much!")

        else:
            ls = ["Try to ask google", "Provide correct command", "Search could not found", "Try to ask correctly",
                  "You are awesome but sry i cannot find results"]
            print(random.choice(ls))
