import speech_recognition as sr
from gtts import gTTS
import warnings
import calendar
import random
import wikipedia
import playsound
from datetime import date
from datetime import datetime
import os
import pyscreenshot as sc
import smtplib
import win32com.client
import webbrowser
import time
import subprocess

def offline(of):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    print(of)
    speaker.Speak(of)
intro="Hello! I'm Harvis virtual assistant created by Harrish Arun...How can i help you!"
offline(intro)
offline("You can do the following things with me:")
uses="You can ask me DATE" \
     "\nYou can ask me TIME" \
     "\nYou can ask me TIME" \
     "\nYou can ask me about any famous persons by using prefix WHO IS" \
     "\nYou can ask me to take a SCREENSHOT" \
     "\nYou can ask me to poen YOUTUBE" \
     "\nYou can search anything in GOOGLE from here itself\nyou can open file EXPLORER by Me\nYou can open NOTEPAD from ME"
print(uses)




warnings.filterwarnings('ignore')
def recordAudio():
    r=sr.Recognizer()
    with sr.Microphone()as source:
        print("say ur query")
        offline("listening....!")

        audio=r.listen(source)

    data=''
    try:
        data=r.recognize_google(audio)
        print("yousaid:"+data)
    except sr.UnknownValueError:

        offline("sorry!,I can't understand...\nType your query here")
        data=input()
    except sr.RequestError as e:
        print('print request result from google speech recognition'+e)
    return data
def assistantresponce(text):
    print(text)
    l=gTTS(text=text,lang='en')
    l.save('assistantresponce.mp3')
    playsound.playsound('assistantresponce.mp3')
    os.remove('assistantresponce.mp3')
def  wakeword(text):
    WAKEWORDS=['hey','ok','hay']
    text=text.lower()
    for phrase in WAKEWORDS:
        if phrase in text:
         return True
    return False


def greeting(text):
    GREETING_INPUTS=["hi",'hello',"ok","hai"]
    GREETING_RESPONSES=['hello','hey there','hola','whatsup']
    for word in text.split():
        if word.lower()in  GREETING_INPUTS:
            return random.choice(GREETING_RESPONSES)+"."
    return ''
def getPerson(text):
    wordList=text.split()

    for i in range(0,len(wordList)):
         if i + 3 <= len(wordList) - 1 and wordList[i].lower() == 'who' and wordList[i+1].lower() == "is":
             return wordList[i+2] +' '+ wordList[i+3]
def offline(of):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    print(of)
    speaker.Speak(of)

def web():

    r3=sr.Recognizer()
    with sr.Microphone() as source:
       audio=r3.listen(source)


while True:
    text=recordAudio()
    responce  =''
    if('wakeword(text)==True'):
        responce=responce+greeting(text)

        offline(responce)

    if('date'in text):
       today=date.today()
       offline("Today's Date is")
       offline(today)

    if("who is" in text):

       person = getPerson(text)
       wiki = wikipedia.summary(person)

       offline(person)
       offline(wiki)

    if('time' in text):
           now=datetime.now()
           current_time=now.strftime("%H:%M:%S:")
           offline(current_time)
    if('screenshot'in text):
        offline('screen shot will be taken in five seconds')

        offline('five')
        time.sleep(1)
        offline('four')
        time.sleep(1)
        offline('three')
        time.sleep(1)
        offline('two')
        time.sleep(1)
        offline('one')
        h = sc.grab()
        h.save('sample.png')
        h.show()
        offline('screen shot taken succesfully')
    if("creator name"  in text):
        offline('I was created by Harrisharun')
    if("mail"in text):
        server=smtplib.SMTP_SSL("smtp.gmail.com",465)
        offline("Enter sender mail id:")
        senderemail=input()
        offline("Enter Your password:")
        password=input()
        offline('enter recivers mail id:')
        recivermail=input()
        offline('Enter the Message Content')
        msg=input()
        server.login(senderemail,password)
        server.sendmail(senderemail,recivermail,msg)
        offline('email has been sent succesfully')
    if("calendar" in text or "today"in text):
         yea=int(input("year:"))
         mont=int(input("month:"))
         print(calendar.month(yea,mont))


    if("YouTube" in text):
        r1 = sr.Recognizer()

        web()
        r2=sr.Recognizer()
        url='https://www.youtube.com/results?search_query='
        with sr.Microphone() as source:
            offline('what you want to search on youtube')
            audio=r1.listen(source)
            try:
                get=r1.recognize_google(audio)
                print(get)
                webbrowser.get().open_new(url+get)
            except:
                offline("sorry!i can't understand")
    if("Google" in text):
        r1 = sr.Recognizer()
        web()
        r2 = sr.Recognizer()
        url = 'https://www.youtube.com/results?search_query='
        with sr.Microphone() as source:
            offline('what you want to search on google')
            audio = r1.listen(source)
            try:
                get = r1.recognize_google(audio)
                print(get)

                webbrowser.open('https://www.google.com/search?safe=active&sxsrf=ALeKk00taNH09RUSRYwr4fqIe3GOiJ23-A%3A1596017389399&source=hp&ei=7UohX5rKFoie4-EPpPq_mAE&q=' +get)
            except:
                 offline("i can't understand")
    if("song" in text):
        offline('media player opened successfully!')
        subprocess.call("C://Program Files (x86)//AIMP3//aimp3.exe")

    if("notepad" in text):
        offline('notepad opened successfully!')
        subprocess.call("notepad.exe")

    if("Explorer" in text):
        offline('file explorer opened successfully!')
        subprocess.call("explorer.exe")

