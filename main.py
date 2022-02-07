import pyttsx3
from googletrans import Translator
import speech_recognition as sr
from docx import Document
import datetime
import os
import smtplib
import ctypes
from tensor import hide

language = 'en'

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
if language == 'en':
    engine.setProperty('voice', voices[1].id)
elif language == 'fr':
    engine.setProperty('voice', voices[0].id)


def speak(audio):
    if language == 'fr':
        audio = translateToFrench(audio)
    engine.say(audio)
    engine.runAndWait()


def translateToFrench(text):
    translator = Translator()
    translation = translator.translate(text, dest="fr")
    return translation.text

def translateToEnglish(text):
    translator = Translator()
    translation = translator.translate(text, dest="en")
    return translation.text


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 8 and hour < 12:
        speak("Good Morning !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon !")

    else:
        speak("Hey!")

    speak("I am your Assistant Vermera")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        if language == 'fr':
            query = r.recognize_google(audio, language='fr')
            query = translateToEnglish(query)
            print('quer' + query)
        else:
            query = r.recognize_google(audio, language='en-in')
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
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # This Function will clean any
    # command before execution of this python file

    clear()
    wishMe()

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('vermera.txt', 'w')
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            file.write(strTime)
            file.write(" :- ")
            file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("vermera.txt", "r")
            print(file.read())
            speak(file.read(6))

        #changing language command
        elif "change" in query:
            if language == 'fr':
                language='en'
            else:
                language='fr'
            speak("Language Changed")
            if language == 'en':
                 engine.setProperty('voice', voices[1].id)
            elif language == 'fr':
                 engine.setProperty('voice', voices[0].id)


        elif "write" in query:

            document = Document()
            document.add_heading('fahd Title', 0)
            document.add_paragraph('hello world im newton world')
            document.save('test.docx')
            os.startfile('test.docx')
        else:
            hide() #call for method

