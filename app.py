from flask import Flask, render_template, request
import openpyxl
import os

app = Flask(__name__)

EXCEL_FILE = "students.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form["name"]
        parent_name = request.form["parent_name"]
        dob = request.form["dob"]
        card_number = request.form["card_number"]
        phone = request.form["phone"]
        bank_name = request.form["bank_name"]

        if not os.path.exists(EXCEL_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["اسم الطالب", "اسم ولي الأمر", "تاريخ الميلاد", "رقم بطاقة السكن", "رقم الهاتف", "اسم المصرف"])
            wb.save(EXCEL_FILE)

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([name, parent_name, dob, card_number, phone, bank_name])
        wb.save(EXCEL_FILE)

        return "تم حفظ المعلومات بنجاح ✅"

    return render_template("form.html")
