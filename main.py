from googletrans import Translator
from docx import Document
from documents import create_psd, create_pfr, create_test_plan
from mailer import send_mail
from questions import get_id
from todo import show_todo, show_history, insert_todo, finish_todo
from openpyxl import load_workbook


def showtodo(mail):
    todo = ""
    df = show_todo(mail)
    df = df.reset_index()
    for index, row in df.iterrows():
        todo = todo + str(index) + "- " + row['todo'] + " --- " + row['time'] + "      \n"
    return todo


def showhistory(mail):
    todo = ""
    df = show_history(mail)
    df = df.reset_index()
    for index, row in df.iterrows():
        todo = todo + str(index) + "- " + row['todo'] + " --- " + row['time'] + "\n"
    return todo


def translateToFrench(text):
    translator = Translator()
    translation = translator.translate(text, dest="fr")
    return translation.text



def takeCommand():
    return "ok"


def email(msg, question, mail):
    if question == "What should I write?":
        send_mail(msg, mail, "Generated Mail")
        return "Mail Generated Has been sent to your address"


def translate(msg, question):
    if question == "what do you want to translate?":
        return translateToFrench(msg)


def add_to_do(msg, question, mail):
    if question == "what is the to do to add?":
        insert_todo(msg, mail)
        return "To do added successfully."
def complete_to_do(msg, question):
    if question == "what is the number of to do to complete?":
        try:
            finish_todo(msg)
        except ValueError:
            print("Invalid Number entered")
            return "number entered not valid"
        return "To do completed successfully."


def bot_functions(query, mail):
    if 'mail' in query:
        return "What should I write?"

    elif "translate" in query:
        return "what do you want to translate?"

    elif "show to do" in query:
        return showtodo(mail)

    elif "add to do" in query:
        return "what is the to do to add?"

    elif "complete to do" in query:
        return "what is the number of to do to complete?"

    elif "history to do" in query:
        return showhistory(mail)
    elif ("psd" in query) or ("developer document" in query):
        document = Document('utils/PSD.docx')
        document.save(mail + "PSD.docx")
        return "PSD in creation, what do you want as a title?"
    elif "client document" in query:
        document = Document('utils/PFR.docx')
        document.save(mail + "PFR.docx")
        return "PFR in creation, what do you want as a title?"
    elif "test document" in query:
        workbook = load_workbook(filename='utils/plan.xlsx')
        workbook.save(mail + "plan.xlsx")
        return "test plan in creation, what is the title of the test?"
    else:
        return "model"  # call tensorflow model


def generate_answer(msg, mail, question):
    question_id = get_id(question)
    if question_id == 1:
        return create_psd(msg, question, mail)
    elif question_id == 2:
        return create_pfr(msg, question, mail)
    elif question_id == 3:
        return create_test_plan(msg, question, mail)
    elif question_id == 4:
        return email(msg, question, mail)
    elif question_id == 5:
        return translate(msg, question)
    elif question_id == 6:
        return add_to_do(msg, question, mail)
    elif question_id == 7:
        return complete_to_do(msg, question)
    else:
        return bot_functions(msg.lower(), mail)


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
