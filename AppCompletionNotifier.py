from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import cv2
import pytesseract
import numpy as np
import time
import re
from pytesseract import Output
from PIL import ImageGrab

from pynput import mouse
from enum   import Enum

class CheckType(Enum):
    PERCENTAGE = 1
    NUMBERCOMPARE = 2

#   ---MOUSE CAPTURING VARIABLES---   #
start_point = None
end_point   = None

capturing   = False
#   -------------------------------   #

#   --EMAIL CREDENTIALS VARIABLES--   #
sender_email    = "__.@__.__"
sender_password = "_________"

receiver_email  = "__.@__.__"

email_provider_host = "smtp-______"
email_provider_port = 465
#   -------------------------------   #

#   --------Capture settings-------   #
typeOfCheck             = CheckType.NUMBERCOMPARE
threshholdForCompletion = 95
amountOfCheck           = 5
timeBetweenChecks       = 10
maxErrorsBeforeStop     = 5

#   -------------------------------   #



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

# Configure Pytesseract path to your installation
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path

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

def check_for_percentage(textInput):
    match = re.search(r"(\d+)%", textInput)
    if match:
        percentage = int(match.group(1))
        print(f"Current progress: {percentage}%")
        return percentage
    return -1

def compare_2_numbers(textInput):
    # Use regex to search for the pattern in the OCR result
    match = re.search(r"Submitting file (\d+) of (\d+) to depot", textInput)
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
            print("something went wrong with trying to read the file")
            breakOut = True
        else:
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

def preprocess_for_ocr(img):
    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    return img

if __name__ == "__main__":
    # Listen for the mouse click
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

    if not capturing and start_point and end_point:
        # Once we have the start and end points, calculate the area
        mouseArea = calculate_area(start_point, end_point)
        print("Selected Area: Left-Top (X, Y), Width, Height")
        print(mouseArea)
        
        # Example usage
        capture_and_detect(mouseArea, typeOfCheck, threshholdForCompletion, amountOfCheck, timeBetweenChecks, maxErrorsBeforeStop)
