from googletrans import Translator
import speech_recognition as sr
from docx import Document
import datetime
import os
from openpyxl import Workbook
import streamlit as st
from streamlit_chat import message as st_message
from todo import show_todo, show_history, insert_todo, finish_todo
from gtts import gTTS
from playsound import playsound
import getpass
from openpyxl import load_workbook

global lang
lang = 'en'

global mail
mail = ''


def speak(audio):
    print(lang)
    if lang == 'fr':
        audio = translateToFrench(audio)
    tts = gTTS(text=audio, lang=lang)
    tts.save("say.mp3")
    playsound("say.mp3")
    os.remove("say.mp3")
    st.session_state.history.append(
        {"message": audio, "is_user": False, "avatar_style": "jdenticon"})


def showtodo():
    df = show_todo()
    df = df.reset_index()
    for index, row in df.iterrows():
        st.session_state.history.append(
            {"message": str(index) + "- " + row['todo'] + " --- " + row['time'], "is_user": False,
             "avatar_style": "jdenticon"})


def showhistory():
    df = show_history()
    df = df.reset_index()
    for index, row in df.iterrows():
        st.session_state.history.append(
            {"message": str(index) + "- " + row['todo'] + " --|-- " + row['status'] + " --|-- " + row['time'],
             "is_user": False, "avatar_style": "jdenticon"})


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

    else:
        speak("Hey ! How Can I help you?")
    speak("please enter your vermeg mail before continuing")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        playsound("google.mp3")  # google micro sound
        r.pause_threshold = 1
        query = "None"
        while query == "None":
            audio = r.listen(source, phrase_time_limit=10)
            try:
                print("Recognizing...")
                if lang == 'fr':
                    query = r.recognize_google(audio, language='fr')
                else:
                    query = r.recognize_google(audio, language='en-in')
                st.session_state.history.append({"message": query, "is_user": True, "avatar_style": "micah"})
                print(f"User said: {query}\n")
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


def modifyDoc(document, key, msg):
    for paragraph in document.paragraphs:
        if key in paragraph.text:
            paragraph.text = msg


def modifyTableDoc(document, key, msg):
    for table in document.tables:
        data = []
    keys = None
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            if key in cell.text:
                cell.text = msg


def fill_test_plan(spreadsheet, desc, status):
    finished = True
    i = 4
    while (finished):
        if spreadsheet["A" + str(i)].value is None:
            spreadsheet["A" + str(i)] = "TSC-" + str(i - 3)
            spreadsheet["B" + str(i)] = desc
            spreadsheet["C" + str(i)] = status
            finished = False
            print(i)
        else:
            i = i + 1


def addTitle(document, title):
    document.add_heading(title, 2)


def addDescription(document, desc):
    document.add_paragraph(desc)


def create_psd():
    speak("PSD in creation")
    document = Document('utils/PSD.docx')
    speak("what do you want as a title?")
    modifyDoc(document, "xtitlex", takeCommand())
    from datetime import datetime
    today = datetime.today().strftime('%Y-%m-%d')
    modifyDoc(document, "xdatex", today)
    modifyTableDoc(document, "xdatex", today)
    # usernames
    modifyTableDoc(document, "xauthorx", getpass.getuser())
    speak("what is the client context?")
    modifyDoc(document, "xclient contextx", takeCommand())
    speak("what is the business context?")
    modifyDoc(document, "xbusiness contextx", takeCommand())
    speak("give me a brief description of the change request ?")
    modifyDoc(document, "xdescriptionx", takeCommand())
    answer = "yes"
    document.add_page_break()
    while (answer != "no"):
        speak("do you want to add new feature description ?")
        answer = takeCommand()
        if (answer != "no"):
            speak("what is the feature's name ?")
            addTitle(document, takeCommand())
            speak("what is the description for this feature ?")
            addDescription(document, takeCommand())

    speak("document is now saved on your desktop, i will open it now")
    path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + '\\PSD.docx'
    document.save(path)
    os.startfile(path)


def create_pfr():
    speak("pfr in creation")
    document = Document('utils/PFR.docx')
    speak("what do you want as a title?")
    modifyDoc(document, "xtitlex", takeCommand())
    from datetime import datetime
    today = datetime.today().strftime('%Y-%m-%d')
    modifyDoc(document, "xdatex", today)
    modifyTableDoc(document, "xdatex", today)
    # usernames
    modifyTableDoc(document, "xauthorx", getpass.getuser())
    speak("what is the aim of the document?")
    modifyDoc(document, "xaimx", takeCommand())
    speak("describe the current behavior?")
    modifyDoc(document, "xcurrent behaviorx", takeCommand())
    speak("what is the proposed solution")
    modifyDoc(document, "xsolutionx", takeCommand())
    answer = "yes"
    while (answer != "no"):
        speak("do you want to add new feature description ?")
        answer = takeCommand()
        if (answer != "no"):
            speak("what is the feature's name ?")
            addTitle(document, takeCommand())
            speak("what is the description for this feature ?")
            addDescription(document, takeCommand())

    speak("document is now saved on your desktop, i will open it now")
    path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + '\\PFR.docx'
    document.save(path)
    os.startfile(path)


def create_test_plan():
    speak("test plan in creation")
    workbook = load_workbook(filename='utils/plan.xlsx')
    spreadsheet = workbook.active
    speak("what is the title of the test?")
    title = takeCommand()
    spreadsheet["B1"] = title
    testCase = True
    while testCase:
        speak("what is the description of the test case?")
        desc = takeCommand()
        speak("what is the actual status of this test case ?")
        status = takeCommand()
        fill_test_plan(spreadsheet, desc, status)
        speak("do you want to add another test case ?")
        answer = takeCommand()
        if "no" in answer:
            testCase = False
    path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + '\\Test Plan -' + title + '.xlsx'
    workbook.save(path)
    os.startfile(path)


def bot_functions(query):
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

    elif "translate" in query:
        speak(translateToFrench(takeCommand()))

    elif "show to do" in query:
        speak("here\'s your to do for the day")
        showtodo()

    elif "add to do" in query:
        speak("name of the to do")
        todo = takeCommand()
        insert_todo(todo)
        speak("to do inserted")
        showtodo()

    elif "complete to do" in query:
        showtodo()
        speak("what is the number of to do to complete")
        number = takeCommand()
        try:
            finish_todo(number)
        except ValueError:
            speak("Invalid Number entered")
        showtodo()

    elif "history to do" in query:
        speak("here's the to do history")
        showhistory()

    elif "developer document" in query:
        create_psd()
    elif "client document" in query:
        create_pfr()
    elif "test document" in query:
        create_test_plan()

    else:
        speak("The tensorflow model is not yet supported")  # call tensorflow model


def bot_functions_fr(query):
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


    elif "traduire" in query:
        speak(translateToEnglish(takeCommand()))

    elif ("mes tâches" or "mes taches") in query:
        speak("here\'s your to do for the day")
        showtodo()

    elif ("ajoute tâche"or "ajoute tache") in query:
        speak("name of the to do")
        todo = takeCommand()
        insert_todo(todo)
        speak("to do inserted")
        showtodo()

    elif ("compléter tâche"or "compléter tache"or "completer tache") in query:
        showtodo()
        speak("what is the number of to do to complete")
        number = takeCommand()
        try:
            finish_todo(number)
        except ValueError:
            speak("Invalid Number entered")
        showtodo()

    elif ("historique tâche" or "historique tache") in query:
        speak("here's the to do history")
        showhistory()

    elif "document développement" in query:
        create_psd()
    elif "document client" in query:
        create_pfr()
    elif "plan de test" in query:
        create_test_plan()

    else:
        speak("The tensorflow model is not yet supported")  # call tensorflow model


def generate_answer():
    user_message = st.session_state.input_text
    user_message = user_message.lower()
    st.session_state.history.append({"message": user_message, "is_user": True, "avatar_style": "micah"})
    st.session_state["input_text"] = ""
    if user_message == "vermera":
        if lang == 'fr':
            bot_functions_fr(takeCommand())
        else:
            bot_functions(takeCommand())
    else:
        if lang == 'fr':
            bot_functions_fr(user_message)
        else:
            bot_functions(user_message)


if __name__ == '__main__':

    isToDoExist = os.path.exists('todo.xlsx')
    if not isToDoExist:
        workbook = Workbook()
        spreadsheet = workbook.active
        spreadsheet["A1"] = "todo"
        spreadsheet["B1"] = "time"
        spreadsheet["C1"] = "status"
        workbook.save(filename='todo.xlsx')

    if "history" not in st.session_state:
        st.session_state.history = []
        wishMe()

    st.title("Vermera Virtual Assistant")
    msg_limit = 10000
    st.text_input("Talk to the bot", key="input_text", on_change=generate_answer)
    for chat in reversed(st.session_state.history):
        st_message(chat['message'], chat['is_user'], chat['avatar_style'], None, str(msg_limit))  # unpacking
        msg_limit = msg_limit - 1
    mail = st.sidebar.text_input("Vermeg mail :", key="mail_input")
    option = st.sidebar.selectbox("", ('English', 'Français'))
    if option == 'English':
        lang = 'en'
        st.sidebar.markdown("# Commands")
        st.sidebar.markdown("This app has a lot of different commands : ")
        st.sidebar.markdown("1 - Vermera : to use the microphone")
        st.sidebar.markdown("2 - show to do : to show your to do\'s")
        st.sidebar.markdown("3 - add to do : to add new to do")
        st.sidebar.markdown("4 - complete to do : to complete a to do")
        st.sidebar.markdown("5 - history to do : to show your to do\'s history")
        st.sidebar.markdown("6 - translate : to translate from english to french")
        st.sidebar.markdown("7 - developer document : to Create PSD.")
        st.sidebar.markdown("8 - client document : to Create PFR.")
        st.sidebar.markdown("9 - test document : to Create a test plan.")
    elif option == 'Français':
        lang = 'fr'
        st.sidebar.markdown("# Commande")
        st.sidebar.markdown("Les commandes de l\'assistant vocale : ")
        st.sidebar.markdown("1 - Vermera : pour utiliser le microphone")
        st.sidebar.markdown("2 - mes taches : pour voir vos taches")
        st.sidebar.markdown("3 - ajoute tache : pour ajouter une nouvelle tache")
        st.sidebar.markdown("4 - compléter tache : pour completer une tache")
        st.sidebar.markdown("5 - historique tache : pour voir l\'historique des taches")
        st.sidebar.markdown("6 - traduire : traduire du français vers l'anglais")
        st.sidebar.markdown("7 - document développement : créer PSD.")
        st.sidebar.markdown("8 - document client : créer PFR.")
        st.sidebar.markdown("9 - plan de test : créer un plan de test.")
    print(mail)
