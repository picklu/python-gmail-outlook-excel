import os
import openpyxl

from datetime import datetime
from decouple import config
from string import Template


wb_folder = config('WORKING_FOLDER')
dt_folder = config('DATA_FOLDER')
wb_name = config('WB_NAME')
wb_path = os.path.join(wb_folder, wb_name)
workbook = openpyxl.load_workbook(wb_path)

from_name = config("FROM_NAME")
from_email = config('EMAIL_ID')
email_pass = config('EMAIL_PASSWORD')
cc_email = config('CC_MAIL_ID')


class Student:

    def __init__(self, department):
        """Student object to store all the information
        of a student of a prticular department.

        Args:
            department (string): Name of the department collected 
            from the name of the corresponding worksheet name.
        """
        self.department = department
        self.id = None
        self.name = None
        self.email = None
        self.mobile = None
        self.paid = None
        self.file_path = None

    def __repr__(self) -> str:
        return f"__{self.name}({self.id}) of {self.department}"


def update_student(student, ws, row):
    """Update student object with information from row of ws

    Args:
        student (Student): instance of Student
        ws (Worksheet): worksheet of a workbook
        row (row): row of a worksheet
    """
    for cell in row:
        match ws.cell(1, cell.column).value.lower():
            case "student id":
                student.id = cell.value
            case "name":
                student.name = cell.value
            case "email":
                student.email = cell.value
            case "mobile":
                student.mobile = cell.value
            case "paid":
                student.paid = cell.value == "Paid"
            case "file name":
                student.file_path = os.path.join(
                    dt_folder, cell.value)


def email_content(recipient):
    """Return content of an email for a recipient.

    Args:
        recipient (string): Name of the recipient to be addressed.
    """
    replaceables = {
        "recipient": recipient,
        "current_datetime": datetime.now(
        ).strftime("%d %B %Y, %H:%M:%S %p")
    }

    with open("template.html", "r") as f:
        content = f.read()

    return Template(content).substitute(replaceables)
