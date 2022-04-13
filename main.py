from googletrans import Translator
import speech_recognition as sr
from docx import Document
import datetime
import os
from openpyxl import Workbook
from todo import show_todo, show_history, insert_todo, finish_todo
from gtts import gTTS
from playsound import playsound
import getpass
from openpyxl import load_workbook
from tkinter import *
import customtkinter  # <- import the CustomTkinter module
from PIL import Image, ImageTk
global lang

lang = 'en'
window = customtkinter.CTk()
BG_GRAY = "#000000"
BG_COLOR = "#000000"
MSG_ENTRY_COLOR = "#000000"
TEXT_COLOR = "#EAECEE"

FONT = "Miriam 11"
FONT_BOLD = "Helvetica 13 bold"
window.title("Vermera Virtual Assistant")
window.resizable(width=False, height=False)
window.configure(width=600, height=550, bg=BG_COLOR)
# head label
head_label = customtkinter.CTkLabel(window, bg=BG_COLOR, fg=TEXT_COLOR,text_font=FONT_BOLD,text="VERMEG", pady=10)
head_label.place(relwidth=1)
# text widget
text_widget = Text(window, width=20, height=2, bg=BG_COLOR, fg=TEXT_COLOR,
                   font=FONT)
text_widget.place(relheight=0.88, relwidth=1, rely=0.1)
text_widget.configure(cursor="arrow", state=DISABLED)
# scroll bar
scrollbar = Scrollbar(text_widget)
scrollbar.place(relheight=1, relx=0.974)
scrollbar.configure(command=text_widget.yview)
# bottom label
bottom_label = Label(window, bg=BG_GRAY, height=40)
bottom_label.place(relwidth=1, rely=0.9)
# message entry box
msg_entry = customtkinter.CTkEntry(bottom_label, bg=MSG_ENTRY_COLOR, fg=TEXT_COLOR)
msg_entry.place(relwidth=0.74, relheight=0.05, rely=0.008, relx=0.011)
msg_entry.focus()
# msg_entry.bind("<Return>", _on_enter_pressed)
# send button
send_image = ImageTk.PhotoImage(Image.open("send.png").resize((20, 20), Image.ANTIALIAS))
send_button = customtkinter.CTkButton(master=bottom_label, image=send_image, text="", width=50, height=50,
                                      corner_radius=10, fg_color="gray10", hover_color="gray25",
                                      command=lambda: _on_enter_pressed(msg_entry,text_widget))
send_button.place(relx=0.76, rely=0.008, relheight=0.05, relwidth=0.11)
micro_image = ImageTk.PhotoImage(Image.open("micro.png").resize((20, 20), Image.ANTIALIAS))
micro_button = customtkinter.CTkButton(master=bottom_label, image=micro_image, text="", width=50, height=50,
                                       corner_radius=10, fg_color="#f54251", hover_color="#ed5562",
                                       command=lambda: micro_on(text_widget))
micro_button.place(relx=0.88, rely=0.008, relheight=0.05, relwidth=0.11)


def speak(audio):
    print(lang)
    if lang == 'fr':
        audio = translateToFrench(audio)
    tts = gTTS(text=audio, lang=lang)
    tts.save("say.mp3")
    playsound("say.mp3")
    os.remove("say.mp3")
    label = customtkinter.CTkLabel(master=text_widget, text=audio, fg_color="#17215e", width=150, height=29,
                                   corner_radius=20, text_font=FONT)
    label.pack(padx=5, pady=5, anchor=W)

def showtodo():
    df = show_todo()
    df = df.reset_index()
    for index, row in df.iterrows():
        print("ok")
      #  st.session_state.history.append({"message": str(index) + "- " + row['todo'] + " --- " + row['time'], "is_user": False, "avatar_style": "jdenticon"})


def showhistory():
    df = show_history()
    df = df.reset_index()
    for index, row in df.iterrows():
        print("ok")


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
    document = Document('PSD.docx')
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
    document = Document('PFR.docx')
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
    workbook = load_workbook(filename='plan.xlsx')
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

    elif ("compléter tâche"or "compléter tache"or "compléter tache") in query:
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
    elif "french":
        lang='en'
        speak("Language changed to french")  # call tensorflow model
    else:
        speak("The tensorflow model is not yet supported")  # call tensorflow model


def generate_answer(msg):
    if lang == 'fr':
        bot_functions_fr(msg)
    else:
        bot_functions(msg)




def run():
    window.mainloop()


def micro_on(text_widget):
    msg = takeCommand()
    _insert_message(msg, text_widget)

def _on_enter_pressed(msg_entry, text_widget):
    msg = msg_entry.get()
    msg_entry.delete(0, END)
    _insert_message(msg, text_widget)
def _insert_message(msg, text_widget):
    if not msg:
        return
    label_user = customtkinter.CTkLabel(master=text_widget, text=msg, fg_color="#6677d9", width=100, height=29,
                                        corner_radius=20, text_font=FONT)
    label_user.pack(padx=20, pady=5, anchor=E)
    generate_answer(msg)
    #label = customtkinter.CTkLabel(master=text_widget, text=generate_answer(msg), fg_color="#17215e", width=150, height=29,
    #                               corner_radius=20, text_font=FONT)
    #label.pack(padx=5, pady=5, anchor=W)

if __name__ == '__main__':

    path = os.getenv('APPDATA') + '/Vermera'
    isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)
    isToDoExist = os.path.exists(path + '\\todo.xlsx')
    if not isToDoExist:
        workbook = Workbook()
        spreadsheet = workbook.active
        spreadsheet["A1"] = "todo"
        spreadsheet["B1"] = "time"
        spreadsheet["C1"] = "status"
        workbook.save(filename=path + '\\todo.xlsx')
    run()

