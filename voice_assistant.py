import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[1].id)

def Speak(audio):
    engine.say(audio)
    engine.runAndWait()

def WishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        Speak("Good Morning Sir")
    elif hour>=12 and hour<18:
        Speak("Good Afternoon Sir")

    else:
        Speak("Good Evening Sir") 
    
    assname = ("Jarvis 1 point o")
    Speak("I am you assistant")
    Speak(assname)


def username():
    Speak("What should I call you sir")
    uname = takeCommand()
    if uname is None or uname.strip() == "None":  # Check if uname is None
        uname = "User"
    Speak(f"Hello Mister {uname}")
    columns = shutil.get_terminal_size().columns
    print("#################".center(columns))
    print(f"Welcome Mr. {uname}".center(columns))
    print("#################".center(columns))
    
    Speak("How can I help you sir")

def takeCommand():
    r = sr.Recognizer()
     
    with sr.Microphone() as source:
         
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing")
        query = r.recognize_google(audio, language = 'en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to recognize your voice")
        return None

    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    server.login('Your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()

if __name__ == "__main__":
    clear = lambda: os.system('cls')
    clear()
    WishMe()
    username()

    while True:
        query = takeCommand()
        if query is None or query.strip() == "None":  # Ensure query is not None
            query = ""  # Set default empty string to prevent errors

        query = query.lower() 

        if "exit" in query or "stop" in query or "quit" in query:
            print("Goodbye! Have a great day! 😊")
            Speak("Goodbye! Have a great day!")
            break

        if 'wikipedia' in query:
            Speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query,sentences = 3)
            Speak("According to wikipedia")
            print(results)
            Speak(results)

        elif 'open youtube' in query:
            Speak("Here you go to youtube\n")
            webbrowser.open("youtube.com")

        elif 'open stackoverflow' in query:
            Speak("Here you go to stack overflow flow.happy coding\n")
            webbrowser.open("stackoverflow.com")

        elif 'open google' in query:
            Speak("Here you go  to google\n")
            webbrowser.open("google.com")
        
       
        elif 'music' in query or 'play music' in query or 'spotify' in query or 'song' in query:
            Speak("Here you go with music")
            webbrowser.open("https://www.youtube.com/watch?v=kPa7bsKwL-c")  


        
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            Speak(f"Sir, the time is {strTime}")
 
        codePath = r"C:\\Users\\JAY\\AppData\\Local\\Programs\\Opera\\launcher.exe"
        if os.path.exists(codePath):
            os.startfile(codePath)
            Speak("Opening Opera browser")

 
        elif 'email to jay' in query:
            try:
                Speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"   
                sendEmail(to, content)
                Speak("Email has been sent !")
            except Exception as e:
                print(e)
                Speak("I am not able to send this email")
 
        elif 'send a mail' in query:
            try:
                Speak("What should I say?")
                content = takeCommand()
                speak("whom should i send")
                to = input()    
                sendEmail(to, content)
                Speak("Email has been sent !")
            except Exception as e:
                print(e)
                Speak("I am not able to send this email")
 
        elif 'how are you' in query:
            Speak("I am fine, Thank you")
            Speak("How are you, Sir")
 
        elif 'fine' in query or "good" in query:
            Speak("It's good to know that your fine")
 
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
 
        elif "change name" in query:
            Speak("What would you like to call me, Sir ")
            assname = takeCommand()
            Speak("Thanks for naming me")
 
        elif "what's your name" in query or "What is your name" in query:
            Speak("My friends call me")
            Speak(assname)
            print("My friends call me", assname)
 
        elif 'exit' in query:
            Speak("Thanks for giving me your time")
            exit()
 
        elif "who made you" in query or "who created you" in query: 
            Speak("I have been created by Gaurav.")
             
        elif 'joke' in query:
            Speak(pyjokes.get_joke())
             
        elif "calculate" in query: 
             
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            Speak("The answer is " + answer) 
 
        elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query) 
 
        elif "who i am" in query:
            Speak("If you talk then definitely your human.")
 
        elif "why you came to world" in query:
            Speak("Thanks to Jay. further It's a secret")
 
        # elif 'power point presentation' in query:
        #     Speak("opening Power Point presentation")
        #     power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
        #     os.startfile(power)
 
        elif 'is love' in query:
            Speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            Speak("I am your virtual assistant created by Gaurav")
 
        elif 'reason for you' in query:
            Speak("I was created as a Minor project by Mister Gaurav ")
 
        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                       0, 
                                                       "Location of wallpaper",
                                                       0)
            Speak("Background changed successfully")
 
        elif 'open bluestack' in query:
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)
 
        elif 'news' in query:
             
            try: 
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                Speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                     
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    Speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                 
                print(str(e))
 
         
        elif 'lock window' in query:
                Speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
        elif 'shutdown system' in query:
                Speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                 
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            Speak("Recycle Bin Recycled")
 
        elif "don't listen" in query or "stop listening" in query:
            Speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
 
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            Speak("User asked to Locate")
            Speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")
 
        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")
 
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "hibernate" in query or "sleep" in query:
            Speak("Hibernating")
            subprocess.call("shutdown / h")
 
        elif "log off" in query or "sign out" in query:
            Speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
 
        elif "write a note" in query:
            Speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            Speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "show note" in query:
            Speak("Showing Notes")
            file = open("jarvis.txt", "r") 
            print(file.read())
            Speak(file.read(6))
 
        elif "update assistant" in query:
            Speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)
             
            with open("Voice.py", "wb") as Pypdf:
                 
                total_length = int(r.headers.get('content-length'))
                 
                for ch in progress.bar(r.iter_content(chunk_size = 2391975),
                                       expected_size =(total_length / 1024) + 1):
                    if ch:
                      Pypdf.write(ch)
                     
        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "jarvis" in query:
             
            wishMe()
            speak("Jarvis 1 point o in your service Mister")
            speak(assname)
 
        elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather 
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            Speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url) 
            x = response.json() 
             
            if x["code"] != "404": 
                y = x["main"] 
                current_temperature = y["temp"] 
                current_pressure = y["pressure"] 
                current_humidiy = y["humidity"] 
                z = x["weather"] 
                weather_description = z[0]["description"] 
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description)) 
             
            else: 
                Speak(" City Not Found ")
             
        elif "send message " in query:
                # You need to create an account on Twilio to use this service
                account_sid = 'Account Sid key'
                auth_token = 'Auth token'
                client = Client(account_sid, auth_token)
 
                message = client.messages \
                                .create(
                                    body = takeCommand(),
                                    from_='Sender No',
                                    to ='Receiver No'
                                )
 
                print(message.sid)
 
        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")
 
        elif "Good Morning" in query:
            Speak("A warm" +query)
            Speak("How are you Mister")
            Speak(assname)
 
        # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:   
            Speak("I'm not sure about, may be you should give me some time")
 
        elif "how are you" in query:
            Speak("I'm fine, glad you me that")
 
        elif "i love you" in query:
            Speak("It's hard to understand")
 
        elif "what is" in query or "who is" in query:
             
            # Use the same API key 
            # that we have generated earlier
            client = wolframalpha.Client("API_ID")
            res = client.query(query)
             
            try:
                print (next(res.results).text)
                Speak (next(res.results).text)
            except StopIteration:
                print ("No results")
