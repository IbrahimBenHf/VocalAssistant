import os

from flask import Flask, render_template, request, jsonify
from openpyxl import Workbook

from chat import get_response
from main import generate_answer, translateToFrench
from mainFR import generate_answer_fr, translateToEnglish

app = Flask(__name__)


@app.get("/")
def index_get():
    return render_template("base.html")


@app.post("/predict")
def predict():
    text = request.get_json().get("message")
    lang = request.get_json().get("language")
    mail = request.get_json().get("mail")
    question = request.get_json().get("question")
    # if "@vermeg.com" not in mail:
    #     message = {"answer": "Vermeg mail required to continue!"}
    #     return jsonify(message)
    if lang == "en":
        response = generate_answer(text, mail,question)
    else:
        response = generate_answer_fr(text, mail,question)

    if response == "model":
        response = get_response(text)
    elif response == "model_fr":
        msg = translateToEnglish(text)
        text = get_response(msg)
        response = translateToFrench(text)
    message = {"answer": response}
    return jsonify(message)


if __name__ == "__main__":
    isToDoExist = os.path.exists('utils/todo.xlsx')
    if not isToDoExist:
        workbook = Workbook()
        spreadsheet = workbook.active
        spreadsheet["A1"] = "todo"
        spreadsheet["B1"] = "time"
        spreadsheet["C1"] = "status"
        spreadsheet["D1"] = "mail"
        workbook.save(filename='utils/todo.xlsx')
    app.run(debug=True)
