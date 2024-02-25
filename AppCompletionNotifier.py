from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import imaplib
from email.parser import BytesParser
from email.policy import default

import cv2
import pytesseract
import numpy as np
import time
import re
from pytesseract import Output
from PIL import ImageGrab

from dataclasses import dataclass, field
from typing import Tuple

from pynput import mouse
from enum   import Enum

# Configure Pytesseract path to your installation
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path

class CheckType(Enum):
    PERCENTAGE = 1
    NUMBERCOMPARE = 2

@dataclass
class Job:
    area: Tuple[int, int, int, int]
    check_Type : CheckType
    percentage : int = 0
    currentCount : int = 0
    completed : bool = field(default=False, init=False)


#   ----------JOB-SYSTEM----------   #

jobs = []

job_info_pattern = re.compile(
    r"ADD_JOB\r\n"
    r"area: (?P<area>\d+, \d+, \d+, \d+)\r\n"
    r"checkType: (?P<check_type>\w+)\r"
)


#   -------------------------------   #


#   ---MOUSE-CAPTURING-VARIABLES---   #

start_point = None
end_point   = None

capturing   = False

#   -------------------------------   #

#   --EMAIL-CREDENTIALS-VARIABLES--   #

sender_email    = "__.@__.__"
sender_password = "_________"

receiver_email  = "__.@__.__"

email_provider_host = "smtp-______"
email_provider_port = 465

email_provider_imap_server = '___________'
email_provider_imap_port   = '___'

#   -------------------------------   #

#   --------Capture-settings-------   #

completionThreshold = 95
update_interval     = 10

#   -------------------------------   #

# JOB HANDLING
# --------------------------------------------------

def add_job(job : Job):
    jobs.append(job)

def update_jobs():
    for job in jobs:
        if not job.completed:
            check_job(job)

            if job.completed == True:
                complete_job(job)

    #removing completed jobs from the list
    #global jobs
    #jobs = [job for job in jobs if not job.completed]

def check_job(job: Job):
    print("Checking job...")

    job.currentCount = job.currentCount + 1
    job.percentage = capture_and_detect_job(job.area, job.check_Type)

    if job.percentage >= completionThreshold:
        print("Job has been completed")
        job.completed = True
    else:
        print(f"Job is {job.percentage}% done.")

def complete_job(job : Job):
    print("Handling job completion")
    # Send email to notify the user
    send_email(f"Task {job.percentage}% Completed", f"Your task is {job.percentage}% done. You may want to check soon!", receiver_email)

# --------------------------------------------------

# COMMANDS
# --------------------------------------------------

def null_command(commandLine):
    print(f"Command '{commandLine}' not recognized")

def status_update(body):
    print("Progress status is on ???%")

def send_screenshot(body):
    print("Sending screenshot...")

def command_numberCompare(body):
    print("running check on 2 numbers")
    # capture_and_detect(checkArea, CheckType.NUMBERCOMPARE, 95, -1, update_interval, errorMax)
    # Start job to compare certain numbers
    # A job will check a saved area, untill it has reached its threshold

def command_percentageCheck(body):
    print("running check on percentage")

def command_crashCheck(body):
    print("running check for application crash")

def command_generateJob(body):
    print("Generating job...")
    match = job_info_pattern.search(body)  # Use .search() to scan through the entire body
    if match:
        info = match.groupdict()
        job_area = tuple(map(int, info["area"].split(", ")))
        
        check_type_str = info['check_type'].upper()  # Ensure uppercase for enum match
        try:
            check_type_enum = CheckType[check_type_str]  # Convert string to CheckType enum
        except KeyError:
            print(f"Check type {check_type_str} is not valid. Job not added.")
            return
        
        new_job = Job(area=job_area, check_Type=check_type_enum)
        add_job(new_job)
        print("Job added successfully.")
    else:
        print("Invalid job info format.")

# --------------------------------------------------


# MAILING
# --------------------------------------------------

def check_mail():
    mail = connect_to_email(email_provider_imap_server, sender_email, sender_password)
    email_ids = search_emails(mail, 'UNSEEN')
    for email_id in email_ids:
        email_message = fetch_and_parse_email(mail, email_id)
        subject, body = get_email_content(email_message)
        parse_email_for_commands(subject, body)
    mail.logout()

def connect_to_email(imap_server, email_address, email_password):
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_address, email_password)
    return mail

def search_emails(mail, criteria='ALL'):
    mail.select('inbox')  # Select the inbox folder
    result, data = mail.search(None, criteria)
    if result == 'OK':
        return data[0].split()
    else:
        return []

def fetch_and_parse_email(mail, email_id):
    result, data = mail.fetch(email_id, '(RFC822)')
    if result == 'OK':
        email_bytes = data[0][1]
        parser = BytesParser(policy=default)
        email_message = parser.parsebytes(email_bytes)
        return email_message
    else:
        return None

def get_email_content(email_message):
    subject = email_message['subject']
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if content_type == "text/plain" and "attachment" not in content_disposition:
                body = part.get_payload(decode=True).decode()  # Decode byte to str
                break
    else:
        body = email_message.get_payload(decode=True).decode()
    
    return subject, body

def parse_email_for_commands(subject, body):
    commands = {
        "STATUS_UPDATE": status_update,
        "SEND_SCREENSHOT": send_screenshot,
        "ADD_JOB": command_generateJob,
    }
    
    if subject != "COMMAND":
        return

    for line in body.splitlines():
        line = line.strip().upper()
        if line in commands:
            commands[line](body)
            break
        else:
            null_command(line)

def send_email(subject, message, to_email):    
    # Setup the MIME
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # Add in the message body
    msg.attach(MIMEText(message, 'plain'))
    
    # Create server object without SSL option
    server = SMTP(email_provider_host, email_provider_port)
    
    # Upgrade the connection to secure
    server.starttls()
    
    # Login to the email server
    server.login(sender_email, sender_password)
    
    # Convert the message to a string and send it
    text = msg.as_string()
    server.sendmail(sender_email, to_email, text)
    
    # Quit the server
    server.quit()

# --------------------------------------------------

def on_click(x, y, button, pressed):
    global start_point, end_point, capturing
    if pressed:
        # When the button is pressed, record the starting point
        start_point = (x, y)
        capturing = True
    else:
        # When the button is released, record the ending point and stop listening
        end_point = (x, y)
        capturing = False
        return False  # Return False to stop the listener

def calculate_area(start_point, end_point):
    # Ensure coordinates are in the correct order
    x1, y1 = start_point
    x2, y2 = end_point
    left = min(x1, x2)
    top = min(y1, y2)
    width = abs(x1 - x2)
    height = abs(y1 - y2)
    
    return (left, top, left + width, top + height)

def check_for_percentage(textInput):
    match = re.search(r"(\d+)%", textInput)
    if match:
        percentage = int(match.group(1))
        print(f"Current progress: {percentage}%")
        return percentage
    return -1

def compare_2_numbers(textInput):
    # Use regex to search for the pattern in the OCR result
    match = re.search(r"(\d+) of (\d+)", textInput)
    if match:
        current_file, total_files = int(match.group(1)), int(match.group(2))
        print(f"Current progress: {current_file} of {total_files}")

        percentage = (current_file / total_files) * 100
        print(f"Progress in percent: {percentage}%")

        return percentage
    return -1

def capture_and_detect(area, checkType = CheckType.NUMBERCOMPARE, completionThreshold = 95, checkAmount = 5, update_interval = 5, errorMax = 5):
    breakOut = False

    keepchecking = False
    if checkAmount == -1:
        keepchecking = True
        checkAmount = 2
    
    errorCount = 0
    loopCount = 0
    while loopCount < checkAmount and not breakOut:
        screenshot = ImageGrab.grab(bbox=area)

        img = np.array(screenshot)
        processed_img = preprocess_for_ocr(img)

        # Apply OCR
        text = pytesseract.image_to_string(processed_img, config='--psm 6')
        #print(text)  # Debug: see what text was recognized
        
        percentage = 0
        if(checkType == CheckType.NUMBERCOMPARE):
            percentage = compare_2_numbers(text)
        elif(checkType == CheckType.PERCENTAGE):
            check_for_percentage(text)

        percentage = round(percentage)
        if percentage >= completionThreshold:
            print("Task completed.")                
            send_email(f"Task {percentage}% Completed", f"Your task is {percentage}% done. You may want to check soon!", receiver_email)
            breakOut = True
        elif percentage == -1:
            print("Progress text not detected, make sure the selected area is correct.")
            errorCount = errorCount + 1
            if errorCount >= errorMax:
                breakOut = True            
        
        if breakOut:
            return

        print("check")
        time.sleep(update_interval)

        if keepchecking == False:
            loopCount = loopCount + 1

def capture_and_detect_job(area, checkType):
    screenshot = ImageGrab.grab(bbox= area)

    img = np.array(screenshot)
    processed_img = preprocess_for_ocr(img)

    text = pytesseract.image_to_string(processed_img, config='--psm 6')
    percentage = 0
    if(checkType == CheckType.NUMBERCOMPARE):
            percentage = compare_2_numbers(text)
    elif(checkType == CheckType.PERCENTAGE):
            percentage = check_for_percentage(text)
    
    return round(percentage)

def preprocess_for_ocr(img):
    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    return img

def gatherMouseArea():
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

    if not capturing and start_point and end_point:
        # Once we have the start and end points, calculate the area
        mouseArea = calculate_area(start_point, end_point)
        print("Selected Area: Left-Top (X, Y), Width, Height")
        print(mouseArea)

def updateLoop():
    check_mail()
    update_jobs()
    time.sleep(update_interval)

if __name__ == "__main__":

    #job_area = (0, 0, 450, 200)
    #job_checkType = CheckType.NUMBERCOMPARE
    #new_job = Job(job_area, job_checkType)
    #add_job(new_job)

    while(True):
        updateLoop()