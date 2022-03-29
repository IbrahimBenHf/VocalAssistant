from googletrans import Translator
import speech_recognition as sr
from docx import Document
import datetime
import os
from tensor import hide
from openpyxl import Workbook
import streamlit as st
from streamlit_chat import message as st_message
from todo import show_todo, show_history, insert_todo, finish_todo
from gtts import gTTS
from playsound import playsound

language = 'en'


# engine = pyttsx3.init('sapi5')
# voices = engine.getProperty('voices')
# if language == 'en':
#     engine.setProperty('voice', voices[1].id)
# elif language == 'fr':
#     engine.setProperty('voice', voices[0].id)
#  engine.runAndWait()

def speak(audio):
    if language == 'fr':
        audio = translateToFrench(audio)
    tts = gTTS(text=audio, lang='en')
    tts.save("say.mp3")
    playsound("say.mp3")
    os.remove("say.mp3")
    st.session_state.history.append(
        {"message": audio, "is_user": False, "avatar_style": "jdenticon"})


def show(msg):
    st.session_state.history.append(
        {"message": msg, "is_user": False, "avatar_style": "jdenticon"})

def change_language():
    if language == 'fr':
        language = 'en'
    else:
        language = 'fr'
    print(language)


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
        speak("Good Morning ! How Can I help you?")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon ! How Can I help you?")

    else:
        speak("Hey! How Can I help you?")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        speak("I'm Listening")
        r.pause_threshold = 1
        query = "None"
        while query == "None":
            audio = r.listen(source, phrase_time_limit=10)
            try:
                print("Recognizing...")
                if language == 'fr':
                    query = r.recognize_google(audio, language='fr')
                    query = translateToEnglish(query)
                else:
                    query = r.recognize_google(audio, language='en-in')
                print(f"User said: {query}\n")
                st.session_state.history.append({"message": query, "is_user": True, "avatar_style": "micah"})
            except Exception as e:
                print(e)
                print("Unable to Recognize your voice.")
                speak("unable to Recognize your voice, can you repeat")

    return query


# def sendEmail(to, subject, content):
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     msg = outlook.CreateItem(0)
#     msg.To = to
#     msg.Subject = subject
#     msg.Body = content
#     msg.Send()
#
#
# def sendMeeting(date, subject, ):
#     outlook = win32com.client.Dispatch("Outlook.Application")
#     appt = outlook.CreateItem(1)  # AppointmentItem
#     appt.Start = "2022-01-02 14:10"  # yyyy-MM-dd hh:mm
#     appt.Subject = "Subject of the meeting"
#     appt.Duration = 60  # In minutes (60 Minutes)
#     appt.Location = "Location Name"
#     appt.MeetingStatus = 1
#     appt.Recipients.Add("ibrahimbenhf@gmail.com")  # Don't end ; as delimiter
#     appt.Save()
#     appt.Send()

def modifyDoc(document, key, msg):
    for paragraph in document.paragraphs:
        if key in paragraph.text:
            paragraph.text = msg

def addTitle(document, title):
    document.add_heading(title, 2)

def addDescription(document, desc):
    document.add_paragraph(desc)

def create_psd():
    speak("PSD in creation")
    document = Document('PSD.docx')
    speak("what do you want as main title?")
    modifyDoc(document, "xmain titlex", takeCommand())
    from datetime import datetime
    modifyDoc(document, "xdatex", datetime.today().strftime('%Y-%m-%d'))
    speak("what is the client context?")
    modifyDoc(document, "xclient contextx", takeCommand())
    speak("what is the business context?")
    modifyDoc(document, "xbusiness contextx", takeCommand())
    speak("give me a brief description of the change request ?")
    modifyDoc(document, "xdescriptionx", takeCommand())
    answer = "yes"
    document.add_page_break()
    while(answer!="no"):
        speak("do you want to add new feature description ?")
        answer=takeCommand()
        if(answer!="no"):
            speak("what is the feature's name ?")
            addTitle(document, takeCommand())
            speak("what is the description for this feature ?")
            addDescription(document, takeCommand())

    speak("document is now saved on your desktop, i will open it now")
    path=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + '\\PSD.docx'
    document.save(path)
    os.startfile(path)


def create_pfr():
    speak("pfr in creation")


def create_test_plan():
    speak("test plan in creation")


def bot_functions(query, language):
    if query == None:
        query = takeCommand().lower()

    # All the commands said by user will be
    # stored here in 'query' and will be
    # converted to lower case for easily
    # recognition of command
    if 'mail' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            speak("what is the subject")
            subject = takeCommand()
            speak("who should i send to")
            to = input()
            # sendEmail(to, subject, content)
            speak("Email has been sent !")
        except Exception as e:
            print(e)
            speak("I am not able to send this email")

    elif 'meeting' in query:
        speak("meeting")  # meeting idea maybe abandoned

    elif 'exit' in query:
        speak("Thanks for giving me your time")
        exit()

    # changing language command
    elif "change" in query:
        speak("change language")
        # if language == 'fr':
        #     language = 'en'
        # else:
        #     language = 'fr'
        # speak("Language Changed")
        # if language == 'en':
        #     engine.setProperty('voice', voices[1].id)
        # elif language == 'fr':
        #     engine.setProperty('voice', voices[0].id)

    elif "speak" in query:
        takeCommand()


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
        show(show_todo())
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

    elif "developer document" in query:
        create_psd()
    elif "client document" in query:
        create_pfr()
    elif "test document" in query:
        create_test_plan()

    else:
        speak("okay")  # call for method


def generate_answer():
    user_message = st.session_state.input_text
    user_message = user_message.lower()
    st.session_state.history.append({"message": user_message, "is_user": True, "avatar_style": "micah"})
    st.session_state["input_text"] = ""
    if user_message == "vermera":
        bot_functions(takeCommand(), language)
    else:
        bot_functions(user_message, language)


if __name__ == '__main__':

    path = os.getenv('APPDATA') + '/Vermera'
    isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)
    if "history" not in st.session_state:
        st.session_state.history = []
        wishMe()

    st.title("Vermera Virtual Assistant")
    msg_limit = 10000
    st.text_input("Talk to the bot", key="input_text", on_change=generate_answer)
    for chat in reversed(st.session_state.history):
        st_message(chat['message'], chat['is_user'], chat['avatar_style'], None, str(msg_limit))  # unpacking
        print(chat['message'])
        msg_limit = msg_limit - 1

    st.sidebar.markdown("# Vermera Chatbot")
    st.sidebar.markdown("An intelligent voice assistant aimed to help employees with their daily acivities.")
    st.sidebar.markdown("You can either say or type what you want to say.")
    st.sidebar.markdown("# Commands")
    st.sidebar.markdown("This app has a lot of differen commands : ")
    st.sidebar.markdown("1 - Mail : to send a mail")
    st.sidebar.markdown("2 - Meeting : to schedule a meeting")
    st.sidebar.markdown("3 - To do : to manage todo's")
    st.sidebar.markdown("4 - Document : to start generating a PSD, PFR or test plan document")
    st.sidebar.markdown("5 - Translate : to start translation service.")
    st.sidebar.markdown("# Choose an App to run")
    option = st.sidebar.selectbox("", ('--  No selected app  --', 'Enter App', 'Window App', 'Exit App'))

# while True:
#   bot_functions(None)
