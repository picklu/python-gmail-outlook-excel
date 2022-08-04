import os
import openpyxl
from datetime import datetime
from gmail import GMail, Message
from decouple import config


workbook_folder = "D:\\"
data_folder = "D:\\term_results"
workbook_name = "spring 2022 results to be sent to.xlsx"
workbook_path = os.path.join(workbook_folder, workbook_name)
workbook = openpyxl.load_workbook(workbook_path)


class Student:

    def __init__(self, department):
        self.department = department
        self.id = None
        self.name = None
        self.file_path = None
        self.email = None
        self.mobile = None
        self.paid = False

    def __repr__(self) -> str:
        return f"__{self.name}({self.id}) of {self.department}"


def send_gmail(mail, student, **mail_property):
    """Send email using gmail. Use your App Password 
    to avoid any error sending Gmail
    """

    mail_property["html"] = f"""
    Dear {student.name},<br><br>
    Please find attahced herewith your final result for Spring 2022.
    <br><br>
    All the best.
    <br><br>
    Dr. Subrata Sarker<br>
    Registrar<br>
    University of Skill Enrichment and Technology<br>
    e-mail: ss.rgstr.uset.edu@gmail.com<br>
    Dispatched at {datetime.now().strftime("%d %B %Y, %H:%M:%S %p")}
    """
    mail_property["attachments"] = [student.file_path]
    msg = Message(**mail_property)
    mail.send(msg)

    print("==>", "mail sent to", student.name, "successfully!")


if __name__ == "__main__":
    ws_names = ["CSE", "English"]
    mail_property = {}
    mail = GMail(
        f"Dr. Subrata Sarker<{config('EMAIL_ID')}>", config('PASSWORD'))
    mail_property["subject"] = 'Final result for Spring 2022'
    mail_property["cc"] = config('CC_MAIL_ID')
    mail_property["text"] = "Final result for Spring 2020 is availabe now"

    for ws_name in ws_names:
        ws = workbook[ws_name]

        for row in ws.iter_rows(min_row=2):
            if row[0].value == None:
                continue

            student = Student(ws_name)

            for cell in row:
                if ws.cell(1, cell.column).value == "Student Id":
                    student.id = cell.value
                if ws.cell(1, cell.column).value == "Name":
                    student.name = cell.value
                if ws.cell(1, cell.column).value == "file name":
                    student.file_path = os.path.join(data_folder, cell.value)
                if ws.cell(1, cell.column).value == "email":
                    student.email = cell.value
                if ws.cell(1, cell.column).value == "Mobile":
                    student.mobile = cell.value
                if ws.cell(1, cell.column).value == "Paid":
                    student.paid = cell.value

            if student.name and student.paid:
                mail_property["to"] = "picklumithu@gmail.com"
                send_gmail(mail, student, **mail_property)
    print("==> Done!")