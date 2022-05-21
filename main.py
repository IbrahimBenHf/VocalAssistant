from googletrans import Translator
from docx import Document
import os
from openpyxl import Workbook
from datetime import datetime
from mailer import send_mail_with_attachment
from questions import get_id
from todo import show_todo, show_history, insert_todo, finish_todo
import getpass
from openpyxl import load_workbook


def showtodo():
    df = show_todo()
    df = df.reset_index()
    for index, row in df.iterrows():
        print(
            {"message": str(index) + "- " + row['todo'] + " --- " + row['time'], "is_user": False,
             "avatar_style": "jdenticon"})


def showhistory():
    df = show_history()
    df = df.reset_index()
    for index, row in df.iterrows():
        print(
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


def takeCommand():
    return "ok"


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
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                if key in cell.text:
                    cell.text = msg


def fill_test_plan_desc(spreadsheet, desc):
    finished = True
    i = 4
    while (finished):
        if spreadsheet["A" + str(i)].value is None:
            spreadsheet["A" + str(i)] = "TSC-" + str(i - 3)
            spreadsheet["B" + str(i)] = desc
            finished = False
        else:
            i = i + 1
def fill_test_plan_status(spreadsheet, status):
    finished = True
    i = 4
    while (finished):
        if spreadsheet["A" + str(i+1)].value is None:
            spreadsheet["C" + str(i)] = status
            finished = False
        else:
            i = i + 1


def addTitle(document, title):
    document.add_heading(title, 2)


def addDescription(document, desc):
    document.add_paragraph(desc)


def create_psd(msg, question, mail):
    filename = mail + "PSD.docx"
    if "PSD in creation, what do you want as a title?" == question:
        document = Document(filename)
        modifyDoc(document, "xtitlex", msg)
        today = datetime.today().strftime('%Y-%m-%d')
        modifyDoc(document, "xdatex", today)
        modifyTableDoc(document, "xdatex", today)
        modifyTableDoc(document, "xauthorx", mail)
        document.save(filename)
        return "what is the client context?"
    elif "what is the client context?" == question:
        document = Document(filename)
        modifyDoc(document, "xclient contextx", msg)
        document.save(filename)
        return "what is the business context?"
    elif "what is the business context?" == question:
        document = Document(filename)
        modifyDoc(document, "xbusiness contextx", msg)
        document.save(filename)
        return "give me a brief description of the change request ?"
    elif "give me a brief description of the change request ?" == question:
        document = Document(filename)
        modifyDoc(document, "Xdescriptionx", msg)
        document.save(filename)
        return "do you want to add new feature description ?"
    elif "do you want to add new feature description ?" == question and ("no" not in msg):
        return "what is the feature's title ?"
    elif "what is the feature's title ?" == question:
        document = Document(filename)
        addTitle(document, msg)
        document.save(filename)
        return "what is the description for this feature ?"
    elif "what is the description for this feature ?" == question:
        document = Document(filename)
        addDescription(document, msg)
        document.save(filename)
        return "do you want to add new feature description ?"
    else:
        send_mail_with_attachment(filename, mail, "Generated PSD File")
        return "PSD is now saved and is gonna be sent to you on mail."


def create_pfr(msg, question, mail):
    filename = mail + "PFR.docx"
    if "PFR in creation, what do you want as a title?" == question:
        document = Document(filename)
        modifyDoc(document, "xtitlex", msg)
        today = datetime.today().strftime('%Y-%m-%d')
        modifyDoc(document, "xdatex", today)
        modifyTableDoc(document, "xdatex", today)
        modifyTableDoc(document, "xauthorx", mail)
        document.save(filename)
        return "what is the aim of the document?"
    elif "what is the aim of the document?" == question:
        document = Document(filename)
        modifyDoc(document, "xaimx", msg)
        document.save(filename)
        return "describe the current behavior?"
    elif "describe the current behavior?" == question:
        document = Document(filename)
        modifyDoc(document, "xcurrent behaviorx", msg)
        document.save(filename)
        return "what is the proposed solution?"
    elif "what is the proposed solution?" == question:
        document = Document(filename)
        modifyDoc(document, "xsolutionx", msg)
        document.save(filename)
        return "do you want to add another new feature ?"
    elif "do you want to add another new feature ?" == question and ("no" not in msg):
        return "what is the new feature's title ?"
    elif "what is the new feature's title ?" == question:
        document = Document(filename)
        addTitle(document, msg)
        document.save(filename)
        return "what is the description ?"
    elif "what is the description ?" == question:
        document = Document(filename)
        addDescription(document, msg)
        document.save(filename)
        return "do you want to add another new feature ?"
    else:
        send_mail_with_attachment(filename, mail, "Generated PFR File")
        return "PFR is now saved and is gonna be sent to you on mail."


def create_test_plan(msg, question, mail):
    workbook_path= mail + "plan.docx"
    workbook = load_workbook(filename=workbook_path)
    spreadsheet = workbook.active

    if "test plan in creation, what is the title of the test?" == question:
        spreadsheet["B1"] = msg
        workbook.save(workbook_path)
        return "what is the description of the test case?"
    elif "what is the description of the test case?" ==question:
        fill_test_plan_desc(spreadsheet,msg)
        workbook.save(workbook_path)
        return "what is the actual status of this test case ?"
    elif "what is the actual status of this test case ?" ==question:
        fill_test_plan_status(spreadsheet,msg)
        workbook.save(workbook_path)
        return "do you want to add another test case ?"
    elif "do you want to add another test case ?" ==question ("no" not in msg):
        return "what is the description of the test case?"
    else:
        send_mail_with_attachment(workbook_path, mail, "Generated Test Plan File")
        return "test plan is now saved and is gonna be sent to you on mail."



def bot_functions(quer, mail):
    query =quer.lower()
    if 'mail' in query:
        try:
            print("What should I say?")
            content = takeCommand()
            print("what is the subject")
            subject = takeCommand()
            print("who should i send to")
            to = input()
            # sendEmail(to, subject, content)
            print("Email has been sent !")
        except Exception as e:
            print(e)
            print("I am not able to send this email")

    elif "translate" in query:
        print(translateToFrench(takeCommand()))

    elif "show to do" in query:
        print("here\'s your to do for the day")
        showtodo()

    elif "add to do" in query:
        print("name of the to do")
        todo = takeCommand()
        insert_todo(todo)
        print("to do inserted")
        showtodo()

    elif "complete to do" in query:
        showtodo()
        print("what is the number of to do to complete")
        number = takeCommand()
        try:
            finish_todo(number)
        except ValueError:
            print("Invalid Number entered")
        showtodo()

    elif "history to do" in query:
        print("here's the to do history")
        showhistory()

    elif ("psd" in query) or ("developer document" in query):
        document = Document('utils/PSD.docx')
        document.save(mail + "PSD.docx")
        return "PSD in creation, what do you want as a title?"
    elif "client document" in query:
        document = Document('utils/PSD.docx')
        document.save(mail + "PSD.docx")
        return "PFR in creation, what do you want as a title?"
    elif "test document" in query:
        workbook = load_workbook(filename='utils/plan.xlsx')
        workbook.save(mail + "plan.docx")
        return "test plan in creation, what is the title of the test?"

    else:
        return "model"  # call tensorflow model


def bot_functions_fr(query):
    if query == None:
        query = takeCommand().lower()

    # All the commands said by user will be
    # stored here in 'query' and will be
    # converted to lower case for easily
    # recognition of command
    if 'mail' in query:
        try:
            print("What should I say?")
            content = takeCommand()
            print("what is the subject")
            subject = takeCommand()
            print("who should i send to")
            to = input()
            # sendEmail(to, subject, content)
            print("Email has been sent !")
        except Exception as e:
            print(e)
            print("I am not able to send this email")


    elif "traduire" in query:
        print(translateToEnglish(takeCommand()))

    elif ("mes tâches" or "mes taches") in query:
        print("here\'s your to do for the day")
        showtodo()

    elif ("ajoute tâche" or "ajoute tache") in query:
        print("name of the to do")
        todo = takeCommand()
        insert_todo(todo)
        print("to do inserted")
        showtodo()

    elif ("compléter tâche" or "compléter tache" or "completer tache") in query:
        showtodo()
        print("what is the number of to do to complete")
        number = takeCommand()
        try:
            finish_todo(number)
        except ValueError:
            print("Invalid Number entered")
        showtodo()

    elif ("historique tâche" or "historique tache") in query:
        print("here's the to do history")
        showhistory()

    elif "document développement" in query:
        create_psd()
    elif "document client" in query:
        create_pfr()
    elif "plan de test" in query:
        create_test_plan()

    else:
        print("model")  # call tensorflow model


def generate_answer(msg, lang, mail, question):
    id = get_id(question)
    if id != 0:
        if id == 1:
            return create_psd(msg, question, mail)
        if id == 2:
            return create_pfr(msg, question, mail)
        if id == 3:
            return create_test_plan(msg, question, mail)
    else:
        if lang == 'fr-FR':
            bot_functions_fr(msg, mail)
        else:
            return bot_functions(msg, mail)

# if option == 'English':
#     lang = 'en'
#     st.sidebar.markdown("# Commands")
#     st.sidebar.markdown("This app has a lot of different commands : ")
#     st.sidebar.markdown("1 - Vermera : to use the microphone")
#     st.sidebar.markdown("2 - show to do : to show your to do\'s")
#     st.sidebar.markdown("3 - add to do : to add new to do")
#     st.sidebar.markdown("4 - complete to do : to complete a to do")
#     st.sidebar.markdown("5 - history to do : to show your to do\'s history")
#     st.sidebar.markdown("6 - translate : to translate from english to french")
#     st.sidebar.markdown("7 - developer document : to Create PSD.")
#     st.sidebar.markdown("8 - client document : to Create PFR.")
#     st.sidebar.markdown("9 - test document : to Create a test plan.")
# elif option == 'Français':
#     lang = 'fr'
#     st.sidebar.markdown("# Commande")
#     st.sidebar.markdown("Les commandes de l\'assistant vocale : ")
#     st.sidebar.markdown("1 - Vermera : pour utiliser le microphone")
#     st.sidebar.markdown("2 - mes taches : pour voir vos taches")
#     st.sidebar.markdown("3 - ajoute tache : pour ajouter une nouvelle tache")
#     st.sidebar.markdown("4 - compléter tache : pour completer une tache")
#     st.sidebar.markdown("5 - historique tache : pour voir l\'historique des taches")
#     st.sidebar.markdown("6 - traduire : traduire du français vers l'anglais")
#     st.sidebar.markdown("7 - document développement : créer PSD.")
#     st.sidebar.markdown("8 - document client : créer PFR.")
#     st.sidebar.markdown("9 - plan de test : créer un plan de test.")
