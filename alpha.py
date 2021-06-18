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
#from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import cv2



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)



def speak(audio):
    engine.say(audio)
    engine.runAndWait()
 

def wishMe():
    hour = int(datetime.datetime.now().hour)

    if hour>= 0 and hour<12:
        speak("Good Morning Sir !")
  

    elif hour>= 12 and hour<18:
        speak("Good Afternoon Sir !")   
  

    else:
        speak("Good Evening Sir !")  

    assname =("alpha 1 point o")
    speak("I am your Assistant")
    speak(assname)
     
 

def usrname():
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Master")
    speak(uname)
    columns = shutil.get_terminal_size().columns
     
##    print("#####################".center(columns))
    print("Welcome Mr.", uname)
##    print("#####################".center(columns))     
    speak("How can i Help you?, Sir")
 
def takeCommand():
     
    r = sr.Recognizer()     
    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
  

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language ='en-in')
        print(f"User said: {query}\n")



    except Exception as e:
        print(e)    
        print("Unable to Recognize your voice.")  
        return "None"

     
    return query

  

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    # Enable low security in gmail
    server.login('your email', 'your gmail password') # your email suppose = xyz@1gmail.com  , your password= 1233455354
    server.sendmail('your email', to, content) #your email suppose = xyz@1gmail.com
    server.close()





if __name__ == '__main__':
    clear = lambda: os.system('cls')
    assname = "Alpha 1 point o"
     
    # This Function will clean any
    # command before execution of this python file

    clear()
    
    wishMe()
    usrname()


    while True:
         
        query = takeCommand().lower()
         
        # All the commands said by user will be 
        # stored here in 'query' and will be
        # converted to lower case for easily 
        # recognition of command

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)


        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")
 

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")
 

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")   



    # for playing Music/Song
        
        elif 'play music' in query or "play song" in query:          
            speak("Here you go with music")
            music_dir = r"C:\\Users\\travi\\Downloads\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))
            


    # for opening certain application
        
        elif 'launch opera' in query:
            speak("opera is opening")
            codePath = r"C:\\Users\\travi\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(codePath)
            
        elif 'launch code' in query:
            speak("poping up code")
            codePath ="C:\\Users\\travi\\AppData\\Local\\Programs\\Microsoft VS Code\\code.exe"
            os.startfile(codePath)

        elif 'launch notepad' in query:
            speak("opening notepad")
            notepad = r"C:\\Program Files\\Notepad++\\notepad++.exe"
            os.startfile(notepad)
            
        elif 'power point presentation' in query:
            speak("opening Power Point presentation")
            power = r"C:\Users\travi\OneDrive\Desktop\project\voice assistant.pptx"
            os.startfile(power)


    # for sending an email to predifined receiver

        elif 'email to suraj' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "travishshaikh30@gmail.com"   
                sendEmail(to, content)
                speak("Email has been sent !")

            except Exception as e:
                print(e)
                speak("I am not able to send this email")

    # for sending an email to a receiver by typing his/her email 

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input("enter mail id of user:")    
                sendEmail(to, content)
                speak("Email has been sent !")

            except Exception as e:
                print(e)
                speak("I am not able to send this email")
            


        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")


        elif 'fine' in query:
            speak("It's good to know that your fine")

            

        elif 'change my name to' in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")
 

        elif "what's your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)


        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit() 

        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Master.")
             

        elif 'joke' in query:
            speak(pyjokes.get_joke())


    # finding answers using wolframalpha API 
             
        elif "calculate" in query:
            app_id = "74QUR6-3P9TPPPP44"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer)


        elif "what is" in query or "who is" in query:
            client = wolframalpha.Client("74QUR6-3P9TPPPP44")
            res = client.query(query)

             

            try:
                print (next(res.results).text)
                speak (next(res.results).text)

            except StopIteration:
                print ("No results")


        elif 'search' in query or 'play' in query:
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query)



    # About alpha
    
        elif "who am i" in query:
            speak("If you talk then definately you are````````````` human.")
 

        elif "why you came to world" in query:
            speak("Thanks to my Masters. I can't say much it's a secret. Sorry!")

 

        elif "who are you" in query:
            speak("I am your virtual assistant created by Master")
 

        elif 'reason for you' in query:
            speak("I was created as a Minor project by my Master ")

        elif "news" in query:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak("Here are some headlines from the Times of India, Happy reading")
 

##        elif 'change background' in query:
##            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
##            speak("Background changed succesfully")




        elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

 
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

                 

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")


        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop me from listening commands, sir")
            a = int(input("time:"))
            print("Now i will not take command for", a," seconds sir" )
            time.sleep(a)
            speak("I am at your service once again sir")
 

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.co.in/maps/place/"+location)
 

##        elif "camera" in query or "take a photo" in query:
##            ec.capture(0, "false", "img.jpg")

    
        # for taking capturing images
        
        elif "camera" in query or "take a photo" in query or "selfie" in query:

              
            videoCaptureObject = cv2.VideoCapture(0)
    
            result = True

            while(result):

                ret,frame = videoCaptureObject.read()
                cv2.imwrite("NewPicture.jpg",frame)
                result = False
            videoCaptureObject.release()
            cv2.destroyAllWindows()


        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

             

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
 

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('alpha.txt', 'w')
            speak("Sir, Should i include date and time")
            ndt = takeCommand()

            if 'yes' in ndt or 'sure' in ndt:
                strTime=datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)

            else:
                file.write(note)




        elif "show note" in query:
            speak("Showing Notes")
            file = open("alpha.txt", "r") 
            print(file.read())
            speak(file.read(6))


            
        elif "alpha" in query:             
            wishMe()
            speak("alpha 1 point o in your service Master")
            speak(assname)






        elif "send message " in query:
            
            # You need to create an account on Twilio to use this service
            # this is a trial account so you can only send message to your own registered number

            account_sid = 'AC333a9e01870347d45d45a1e958c0fab7'
            auth_token = '664d6b5e1ba83326c41f8532ee9c2e69'
            client = Client(account_sid, auth_token)
            message = client.messages.create(body = takeCommand(),
                                    from_='+12567871865',
                                    to ='=917977013779')
            print(message.sid)

 

        elif "good morning" in query:
            speak("A warm " +query)
            speak("How are you Master")
            speak(assname)


        elif "good afternoon" in query:
            speak("A warm " +query)
            speak("How are you Master")
            speak(assname)


        elif "good evening" in query:
            speak("A warm " +query)
            speak("How are you Master")
            speak(assname)
 

        # some stupid questions asked by some stupid people
        
        elif "will you be my gf" in query or "will you be my bf" in query:   
            speak("I'm not sure about that, may be you should give me some time")
 

        elif "are you ok" in query:
            speak("I'm fine, glad you asked me that")
 

        elif "i love you" in query:
            speak("It's hard to understand")

 

        elif "open facebook" in query:
            webbrowser.open_new_tab("https://www.facebook.com")
            speak("Facebook is open now")


        elif "open gmail" in query :
             webbrowser.open_new_tab("https://www.gmail.com")
             speak("Gmail is open now")

        
        elif "time" in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"the time is {strTime}")

        elif "search" in query or "get me the" in query:
            query = query.replace("search", "")
            webbrowser.open_new_tab(query)
            speak("Here are the results you wanted")



        elif "cricket score" in query:
            cricket = webbrowser.open_new_tab("https://www.cricbuzz.com/cricket-match/live-scores")
            speak("Here is the scores for Cricket")
            time.sleep(3)


        elif "football score" in query:
            football = webbrowser.open_new_tab("https://www.livescore.com/en/")
            speak("Here are the scores for Football")
            time.sleep(3)


        elif "hockey score" in query:
            hockey = webbrowser.open_new_tab("https://www.livescore.com/en/hockey/")
            speak("Here are the scores for Hockey")
            time.sleep(3)


        elif "basketball score" in query:
            basketball = webbrowser.open_new_tab("https://www.livescore.com/en/basketball/")
            speak("Here are the scores for Basketball")
            time.sleep(3)


        elif "tennis score" in query:
            tennis = webbrowser.open_new_tab("https://www.livescore.com/en/tennis/")
            speak("Here are the scores for Tennis")
            time.sleep(3)

            
        elif "weather" in query:
                api_key = "1caa7608c1b271c9cb8261b0788a486a"
                base_url = "https://api.openweathermap.org/data/2.5/weather?"
                speak("what is the city name")
                city_name = takeCommand()
                complete_url = base_url+"appid="+api_key+"&q="+city_name
                response = requests.get(complete_url)
                x = response.json()
                if x["cod"]!= "404":
                    y=x["main"]
                    current_temperature = y["temp"]
                    current_humidity = y["humidity"]
                    z = x["weather"]
                    weather_description = z[0]["description"]
                    speak(" Temperature in kelvin unit is " +
                      str(current_temperature) +
                      "\n humidity in percentage is " +
                      str(current_humidity) +
                      "\n description " +
                      str(weather_description))
                  
                print(" Temperature in kelvin unit is " +
                      str(current_temperature) +
                      "\n Humidity in percentage is " +
                      str(current_humidity) +
                      "\n Description: " +
                      str(weather_description))
        
        else:
            speak("Sorry sir, I can't do that")
