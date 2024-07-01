import win32com.client as win32
import os
import re
from datetime import datetime
import sys
import pathlib


def get_rfq_info(file_path: str) -> dict:
    # Define the regex pattern
    pattern = r"(RFQ \d+) Quote (\w+)\.pdf"

    # Search for the pattern in the file path
    match = re.search(pattern, file_path)

    if not match:
        return sys.exit("Vendor Attachment Empty")

    rfq_vendor_info = {
        "rfq_number": match.group(0).split(' Quote')[0],
        "rfq_recipient": match.group(2),
    }

    return rfq_vendor_info


def get_greeting() -> str:
    """
    Get the appropriate greeting based on the current time.

    Returns:
    - str: "Good Morning" if the current time is before 12 PM, otherwise "Good Afternoon".
    """
    current_hour = datetime.now().hour
    greeting = "Good Morning," if current_hour < 12 else "Good Afternoon,"

    return f"{greeting} \n\nPlease sent Quote with attachment"


recipient_mapping = {
    'Graybar': "JeffreyAbraham27@gmail.com",
    'Eaton': "Jabraham4849@eagle.fgcu.edu",
    'RS': "JeffreyAbraham@gmail.com; Jabraham4849@eagle.fgcu.edu; WesterlyRocket947@gmail.com",
}


def send_email(rfq_file_path: str):
    info = get_rfq_info(rfq_file_path)
    subject = info['rfq_number']
    vendor = info['rfq_recipient']

    # Open Outlook
    outlook = win32.Dispatch("Outlook.Application")
    outlook_namespace = outlook.GetNamespace("MAPI")

    file_path = pathlib.Path(rfq_file_path)
    file_path_absolute = str(file_path.absolute())

    mailItem = outlook.CreateItem(0)
    mailItem.BodyFormat = 2
    mailItem.HTMLBody = "<HTML Markup>"
    mailItem.Attachments.Add(file_path_absolute)

    mailItem.Subject = subject
    mailItem.Body = get_greeting()
    mailItem.To = recipient_mapping[vendor]
    mailItem.Cc = "jabraham4849@eagle.fgcu.edu"

    mailItem.Save()


def retrieve_attachments():
    RFQ_ATTACHMENT_PATH = r"C:\Users\Jeffr\PycharmProjects\RFQ_Automation\RFQ_attachment.txt"
    try:
        # get the size of file
        file_size = os.path.getsize(RFQ_ATTACHMENT_PATH)

        # if file size is 0, it is empty
        if file_size == 0:
            sys.exit("No RFQ Attachment")

        file = open(RFQ_ATTACHMENT_PATH, "r")
        lines = file.readlines()
        for file_path in lines:
            send_email(file_path)

    # if file does not exist, then exception occurs
    except FileNotFoundError as e:
        sys.exit("File Not Found")


retrieve_attachments()
