import os

from gmail import GMail, Message

from helper import (dt_folder, workbook, Student,
                    update_student, from_email, email_pass,
                    cc_email, email_content)


def send_gmail(mail, student):
    """Send email using gmail. Use your App Password 
    to avoid any error sending Gmail

    Args:
        mail (object): from gmail api
        student (object): info of a student  
    """

    mail_property = {}
    mail_property["subject"] = f"Final result for Summer 2022"
    mail_property["to"] = student.email
    mail_property["cc"] = cc_email
    mail_property["text"] = "The final results for Summer 2022 is available now"
    mail_property["html"] = email_content(student.name)
    mail_property["attachments"] = [student.file_path]
    msg = Message(**mail_property)
    mail.send(msg)

    print(f"==> Mail sent to {student.name}<{student.email}> [{student.id}]")


def get_students(ws_names):
    """returns students with information from worksheets

    Args:
        ws_names (list): list of worksheet names

    Returns:
        list: list of students of type Student
    """
    students = []
    for ws_name in ws_names:
        ws = workbook[ws_name]

        for row in ws.iter_rows(min_row=2):
            if row[0].value == None:
                continue

            student = Student(ws_name)
            update_student(student, ws, row)
            students.append(student)
    return students


if __name__ == "__main__":

    mail = GMail(
        f"Dr. Subrata Sarker<{from_email}>", email_pass)

    ws_names = ["TEST"]

    students = get_students(ws_names)
    for student in students:
        if student.name and student.paid:
            print(
                f"==> Mail to be sent to {student.name}<{student.email}> [{student.id}]")
            send_gmail(mail, student)
        else:
            print(
                f"==O Mail not to be sent to {student.name}<{student.email}> [{student.id}]")
    print("==> Done!")
