from datetime import datetime
from pathlib import Path
import re

from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

app = Flask(__name__)
REGISTER_FILE = Path(__file__).resolve().parent / "Register.xlsx"
EMAIL_PATTERN = re.compile(
    r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+"
)
HEADERS = ["Овог", "Нэр", "Утасны дугаар", "Имэйл", "Бүртгэсэн огноо"]


def save_registration(values):
    if REGISTER_FILE.exists():
        workbook = load_workbook(REGISTER_FILE)
        worksheet = workbook["Бүртгэл"] if "Бүртгэл" in workbook.sheetnames else workbook.active
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Бүртгэл"

    if worksheet.max_row == 1 and worksheet["A1"].value is None:
        worksheet.append(HEADERS)
        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="173B4D")
            cell.alignment = Alignment(vertical="center")

    row = worksheet.max_row + 1
    for column, value in enumerate(
        [*values, datetime.now().strftime("%Y-%m-%d %H:%M:%S")], start=1
    ):
        cell = worksheet.cell(row=row, column=column, value=value)
        cell.data_type = "s"
        cell.number_format = "@"

    for column, width in enumerate([22, 22, 20, 34, 22], start=1):
        worksheet.column_dimensions[chr(64 + column)].width = width
    workbook.save(REGISTER_FILE)


@app.route("/", methods=["GET", "POST"])
def register():
    values = {"surname": "", "given_name": "", "phone": "", "email": ""}
    error = None

    if request.method == "POST":
        values = {
            key: request.form.get(key, "").strip()
            for key in values
        }
        if not all(values.values()):
            error = "Бүх талбарыг бөглөнө үү."
        elif not EMAIL_PATTERN.fullmatch(values["email"]):
            error = "Имэйл хаягаа name@example.com форматаар оруулна уу."
        else:
            try:
                save_registration(
                    [values["surname"], values["given_name"], values["phone"], values["email"]]
                )
            except OSError:
                app.logger.exception("Бүртгэлийн Excel файлыг хадгалж чадсангүй.")
                error = "Бүртгэлийг хадгалахад алдаа гарлаа. Дахин оролдоно уу."
            else:
                return render_template("index.html", values=values, success=True)

    return render_template("index.html", values=values, error=error, success=False)


if __name__ == "__main__":
    app.run(debug=True)
