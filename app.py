import os

from flask import Flask, render_template, request, jsonify
from openpyxl import Workbook

from chat import get_response
from main import generate_answer

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
    response = generate_answer(text, lang, mail,question)
    if response == "model":
        response = get_response(text)
    message = {"answer": response}
    return jsonify(message)


if __name__ == "__main__":
    isToDoExist = os.path.exists('todo.xlsx')
    if not isToDoExist:
        workbook = Workbook()
        spreadsheet = workbook.active
        spreadsheet["A1"] = "todo"
        spreadsheet["B1"] = "time"
        spreadsheet["C1"] = "status"
        workbook.save(filename='todo.xlsx')
    app.run(debug=True)
