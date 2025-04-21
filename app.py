from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

DATA_FILE = "students_data.xlsx"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/submit", methods=["POST"])
def submit():
    student_name = request.form["student_name"]
    guardian_name = request.form["guardian_name"]
    birth_date = request.form["birth_date"]
    card_number = request.form["card_number"]
    phone_number = request.form["phone_number"]
    bank_name = request.form["bank_name"]

    new_data = {
        "اسم الطالب": student_name,
        "اسم ولي الأمر": guardian_name,
        "تاريخ الميلاد": birth_date,
        "رقم بطاقة السكن": card_number,
        "رقم هاتف ولي الأمر": phone_number,
        "اسم المصرف": bank_name,
        "تاريخ الإضافة": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        df = df.append(new_data, ignore_index=True)
    else:
        df = pd.DataFrame([new_data])

    df.to_excel(DATA_FILE, index=False)

    return send_file(DATA_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
