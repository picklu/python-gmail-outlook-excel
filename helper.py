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

    try:
        tmpl = open("template.html", "r").read()
        content = Template(tmpl)
    except:
        content = Template("")
    finally:
        return content.substitute(replaceables)
