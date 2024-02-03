import pyttsx3
import datetime
import speech_recognition as sr
import requests
import json
import wikipedia
import webbrowser
import pywhatkit
import os
import random
import smtplib
import time



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print (voices[2].id)
engine.setProperty('voice',  voices[0].id)
engine.setProperty('rate',180)

def tell (str):
    from win32com.client import Dispatch
    tell = Dispatch("SAPI.SpVoice")
    tell.Speak(str)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def greet() :
    time = int (datetime.datetime.now().hour)
    if time>=0 and time<12 :
        speak("Good morning sir !")
    elif time>=12 and time<18 :
        speak("Good afternoon sir !")
    else:
        speak("Good evening sir !")

    speak("what can i do for u")

def takecommand():
    t = sr.Recognizer()
    with sr.Microphone() as source:
        print("listening......")
        t.energy_threshold = 1500
        t.pause_threshold = 1
        audio = t.listen(source)

    try:
        print("recorgnnizing........")
        query = t.recognize_google(audio, language='en-in')
        print("you said :", query)

    except Exception as e:
        return "None"
    return query


def reminder():

    speak("What shall I remind you about?")
    text = str(takecommand())
    speak("In how many minutes?")
    local_time =takecommand()
    local_time = local_time.replace("in", "")
    local_time = local_time.replace("mutes", "")
    local_time = int(local_time) * 60
    time.sleep(int(local_time))
    speak("reminder is about")
    speak("this is the time to ")
    speak(text)
    print(text)



def sendemail() :
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login("try162598", "123578946")
    speak('please enter email id ')
    to = input()
    speak("what is in the message sir")
    message = takecommand()
    s.sendmail("try162598",to , message)
    s.quit()
    speak("e-mail send successfully")
    speak("something else sir")


def news() :
    tell("News for today.. Lets begin")
    url = "https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=da347bb0610f4e279098768fc3f21f15"
    news = requests.get(url).text
    news_dict = json.loads(news)
    arts = news_dict['articles']
    for article in arts:
        speak(article['title'])
        print(article['title'])
        tell("Moving on to the next news..Listen Carefully")

    speak("Thanks for listening...")

def sum () :
    speak("Please give me first number sir")
    n1 =takecommand()
    speak("Tell me secound number")
    n2 =takecommand()
    result = (int(n1) +int(n2))
    speak("sum of both",)
    speak(result)

if __name__ == '__main__':
    speak(" This is Minor Project")
    print("This is Minor Project")
    print("150516")
    speak("Created by RamChandra Nargawee ")
    print("Created by RamChandra Nargave ")
    speak("hello")
    speak("I am Your virtual assistant")
    greet()

while True:

    query = takecommand().lower()

    if 'open wikipedia' in query :
        webbrowser.open("wikipedia.com")

    elif 'in' and 'wikipedia' in query:
        speak('Searching wikipedia...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=4)
        print(results)
        speak(results)

    elif 'add' in query:
        sum()

    elif 'news' in query:
        news()

    elif 'write' and 'text' in query:
        tell("please tell what you want to text")
        takecommand()

    elif 'open youtube' in query :
        webbrowser.open("youtube.com")

    elif 'open google' in query :
        webbrowser.open("google.com")

    elif 'open stackoverflow ' in query :
        webbrowser.open("stackoverflow.com")

    elif 'open facebook' in query :
        webbrowser.open("facebook.com")

    elif 'open whatsapp' in query :
        webbrowser.open("web.whatsapp.com")

    elif 'open geeksforgeeks' in query :
        webbrowser.open("geeksforgeeks.com")



    elif 'play' and 'youtube' in query :
        speak("sir what can i play on youtube")
        y = takecommand()
        pywhatkit.playonyt(y)

    elif 'send message' in query :
       speak("tell me the contract number sir")
       no =takecommand()
       cno = str('+91') + str(no)
       print(cno)
       speak("whats the message sir")
       msg = takecommand()
       hor = int(datetime.datetime.now().hour)
       m = int(datetime.datetime.now().minute)
       min = int(m) +int (2)
       pywhatkit.sendwhatmsg(cno,msg,hor,min)

    elif 'sleep' in query :
        time = int(datetime.datetime.now().hour)

        if time >= 6 and time < 17 :
            speak("Have a nice day sir !")

        elif time >= 17 and time < 20:
            speak("Have a nice evening sir !")
        else:
            speak("good night sir !")
        exit()

    elif 'wake up' in query :
        speak("Always for you sir ! ")

    # elif 'how to' in query :
    #     from pywikihow import search_wikihow
    #     speak("Here we go")
    #     v = 1 #maximum value
    #     v2 = search_wikihow(query, v)
    #     assert len(v2) == 1
    #     v2 [0].print()
    #     speak(v2[0].summary)

    elif 'play music' in query :
        directory = 'H:\\music'
        songs =os.listdir(directory)
        #print(songs)
        r = random.randrange( 1,8 )
        os.startfile(os.path.join(directory, songs[r]))

    elif 'play song' in query:
        directory = 'H:\\music'
        songs = os.listdir(directory)
        # print(songs)
        r = random.randrange(1, 8)
        os.startfile(os.path.join(directory, songs[r]))

    elif 'shut down' in query :
        speak("by sir meat you again")
        os.system("shutdown /s")

    elif 'time' in query:
        time = datetime.datetime.now().strftime("%H:%M:%S")
        speak("Sir the time is" )
        speak(time)

    elif 'how are you' in query:
        speak("I am fine sir, what about you")

    elif 'send mail' in query:
        sendemail()

    elif 'reminder' in query :
        reminder()

    elif 'alarm' in query :
        reminder()

    elif 'open chrome' in query :
        codepath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        speak("crome is opening")
        os.startfile(codepath)

    elif 'open notepad plus plus' in query :
        speak("notepad ++ is opening")
        codepath = "C:\\Program Files\\Notepad++\\notepad++.exe"
        os.startfile(codepath)
    elif 'open pycharm' in query :
        speak("pycham is opening")
        codepath = "C:\Program Files\JetBrains\PyCharm Community Edition 2021.2.3\bin\pycharm64.exe"
        os.startfile(codepath)
    else:
        speak("")
