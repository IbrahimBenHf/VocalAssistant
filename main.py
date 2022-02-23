import pyttsx3
from googletrans import Translator
import speech_recognition as sr
from docx import Document
import datetime
import os
import ctypes
import win32com.client
from tensor import hide
from openpyxl import Workbook

from todo import show_todo, show_history, insert_todo, finish_todo

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


def sendEmail(to, subject, content):
    outlook = win32com.client.Dispatch("Outlook.Application")
    msg = outlook.CreateItem(0)
    msg.To = to
    msg.Subject = subject
    msg.Body = content
    msg.Send()


def sendMeeting(date, subject, ):
    outlook = win32com.client.Dispatch("Outlook.Application")
    appt = outlook.CreateItem(1)  # AppointmentItem
    appt.Start = "2022-01-02 14:10"  # yyyy-MM-dd hh:mm
    appt.Subject = "Subject of the meeting"
    appt.Duration = 60  # In minutes (60 Minutes)
    appt.Location = "Location Name"
    appt.MeetingStatus = 1
    appt.Recipients.Add("ibrahimbenhf@gmail.com")  # Don't end ; as delimiter
    appt.Save()
    appt.Send()

def create_psd():
    speak("PSD in creation")

def create_pfr():
    speak("pfr in creation")

def create_test_plan():
    speak("test plan in creation")


if __name__ == '__main__':

    path = os.getenv('APPDATA') + '/Vermera'
    isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)

    clear = lambda: os.system('cls')

    # This Function will clean any command before execution of this python file
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
                speak("what is the subject")
                subject = takeCommand()
                speak("who should i send to")
                to = input()
                sendEmail(to, subject, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'meeting' in query:
            speak("meeting")  # meeting idea maybe abandoned

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "note" in query:
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

        # changing language command
        elif "change" in query:
            if language == 'fr':
                language = 'en'
            else:
                language = 'fr'
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

        elif "to do" in query:
            isToDoExist = os.path.exists(path + 'todo.xlsx')
            if not isToDoExist:
                workbook = Workbook()
                spreadsheet = workbook.active
                spreadsheet["A1"] = "todo"
                spreadsheet["B1"] = "time"
                spreadsheet["C1"] = "status"
                workbook.save(filename="todo.xlsx")

            speak("here\'s your todos for the day")
            show_todo()
            speak("do you want to add new todos for the day")
            response = takeCommand()
            if response == "yes":
                speak("name of the todo")
                todo = takeCommand()
                insert_todo(todo)
                speak("todo inserted")
                show_todo()
            speak("do you want to complete a todo")
            response = takeCommand()
            if response == "yes":
                show_todo()
                speak("what is the number of todo to complete")
                number = takeCommand()
                try:
                    finish_todo(number)
                except ValueError:
                    speak("Invalid Number entered")
                show_todo()
            speak("do you want to see all todo history")
            response = takeCommand()
            if response == "yes":
                show_history()

        elif "document" in query:
            speak("Do you want to create a PSD ?")
            response = takeCommand()
            if response == "yes":
                create_psd()
            speak("Do you want to create a PFR ?")
            response = takeCommand()
            if response == "yes":
                create_pfr()
            speak("Do you want to create a test plan ?")
            response = takeCommand()
            if response == "yes":
                create_test_plan()


        else:
            hide()  # call for method
