#Main Project. PLEASE DO NOT EDIT IF YOU DON'T UNDERSTAND
#Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By

#Python Library Imports
import time
import pymem
import os
import pyautogui
import pygetwindow as gw
import win32clipboard
 
#User information. These are all things the user needs to edit. Just open the file and edit!!
from user_information import username
from user_information import email #This import is to import a file, and withing that file importing a variable called "email"
from user_information import password #This import is to import a file, and withing that file importing a variable called "password"
from user_information import to #This import is to import a file, and within that file - import the To section in Gmail
from user_information import subject_ifapprunning #This import is what Selenium will type in for Subject (If the app is running)
from user_information import subject_ifappNOTrunning #This import is what Selenium will type in for Subject (If the app is NOT running)
from user_information import message_body_ifapprunning #This import is what Selenium will type in for Message Body (If the app is running)
from user_information import message_body_ifappNOTrunning #This import is what Selenium will type in for Message Body (If the app is running)
from user_information import password_veri #Password Verification!
from user_information import application_monitor #Application to monitor
from user_information import path_to_app #This variable is the application location!
from user_information import path_to_screenshot #This import is where the screenshots go!
from user_information import file_name #This import is for the file name
from user_information import user_time #This import is to see how long between email checks!
from user_information import name_of_screenshot
from user_information import taskbar_name

def initialization():
    """ This function is the initialization of Selenium. This can also open up another browser if need be. """
    #Selenium Initialization
    global driver
    driver = webdriver.Chrome()
    driver.get("https://www.gmail.com")
    driver.fullscreen_window()

#Functions for the Selenium Elements
def id(id, click=False, send_keys=""):
    """ This function is to find the ID element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.ID, id)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def name(name, click=False, send_keys=""):
    """ This function is to find the NAME element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.NAME, name)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def xpath(xpathEL, click=False, send_keys=""):
    """ This function is to find the XPATH element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.XPATH, xpathEL)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def link_text(link_text, click=False, send_keys=""):
    """ This function is to find the LINK_TEXT element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.LINK_TEXT, link_text)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def partial_link_text(partial_link_text, click=False, send_keys=""):
    """ This function is to find the PARTIAL_LINK_TEXT element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.PARTIAL_LINK_TEXT, partial_link_text)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def tag_name(tag_name, click=False, send_keys=""):
    """ This function is to find the TAG_NAME element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.TAG_NAME, tag_name)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def class_name(class_name, click=False, send_keys=""):
    """ This function is to find the CLASS_NAME element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.CLASS_NAME, class_name)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

def css_selector(css_selector, click=False, send_keys=""):
    """ This function is to find the CSS_SELECTOR element! """
    #Navigation variable for driver.element
    time.sleep(5)
    navigation = driver.find_element(By.CSS_SELECTOR, css_selector)

    #If statement to see if user wants to send keys or not
    if send_keys == "":
        None
    elif send_keys != "":
        navigation.send_keys(send_keys)

    #If statement to see if user wants to click or not
    if click == True:
        navigation.click()

#Error handling (This will send an email saying how an error occured)
def error_email():
    """ This function is sending an email if the program has an error! This will also
        offer the user a chance to restart the program. """
    initialization()
    #Navigating the Gmail Website
    #Clicking "Email or Phone" section + inputting keys(email) for it
    xpath("//input[@id='identifierId']", True, email)
    #Clicking "Next" 
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Password" section + inputting keys for it
    xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
    #Clicking "Next"
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Compose"
    xpath("//div[@role='button'][normalize-space()='Compose']", True)
    #Clicking "To"
    xpath("//input[@role='combobox']", True, to)
    #Clicking "Email"
    xpath("//input[@placeholder='Subject']", True, "App_ERROR")
    #Clicking "Message Body"
    xpath("//div[@aria-label='Message Body']", True, "There was an unknown error running the app. Applicatin will now restart")
    #Clicking "Send"
    xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
    time.sleep(60)
    application()


def reaffirmation_email():
    """ This function is if the User doesn't know the format of the reply back email
        this will email them again saying, 'Error, the format must be..... """
    initialization()
    #Navigating the Gmail Website
    #Clicking "Email or Phone" section + inputting keys(email) for it
    xpath("//input[@id='identifierId']", True, email)
    #Clicking "Next" 
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Password" section + inputting keys for it
    xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
    #Clicking "Next"
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Compose"
    xpath("//div[@role='button'][normalize-space()='Compose']", True)
    #Clicking "To"
    xpath("//input[@role='combobox']", True, to)
    #Clicking "Email"
    xpath("//input[@placeholder='Subject']", True, "Reply Error")
    #Clicking "Message Body"
    xpath("//div[@aria-label='Message Body']", True, "There was an error reading the reply of the email. This could be cause by invalid format, or even wrong password. We will restart the program for you automatically!")
    #Clicking "Send"
    xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
    time.sleep(60)

    time.sleep(2)
    application()

def screenshot_pyautogui():
    """ This function grabs the screenshot saved. """
    time.sleep(5)
    #Searching in the location in big search bar
    pyautogui.keyDown("ctrl")
    time.sleep(1)
    pyautogui.press("l")
    time.sleep(1)
    pyautogui.press("ctrl")
    time.sleep(1)
    pyautogui.write(path_to_screenshot)
    time.sleep(1)
    pyautogui.hotkey()
    time.sleep(1)
    #Searching for screenshot.png with small bar
    pyautogui.keyDown("ctrl")
    time.sleep(1)
    pyautogui.press("f")
    time.sleep(1)
    pyautogui.press("ctrl")
    time.sleep(1)
    pyautogui.write(name_of_screenshot)
    time.sleep(1)
    pyautogui.press("down")
    time.sleep(1)
    pyautogui.press("down")
    time.sleep(1)
    pyautogui.press("up")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)

#5 If user says Yes or No to screenshot (This needs a time.sleep function then it restarts the program!)
def Yes_Email_Screeshot():
    """ This function is for when the user replies with 'NO'. It send another email 
        asking if they want to open it! """
    initialization()
    #Navigating the Gmail Website
    #Clicking "Email or Phone" section + inputting keys(email) for it
    xpath("//input[@id='identifierId']", True, email)
    #Clicking "Next" 
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Password" section + inputting keys for it
    xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
    #Clicking "Next"
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Compose"
    xpath("//div[@role='button'][normalize-space()='Compose']", True)
    #Clicking "To"
    xpath("//input[@role='combobox']", True, to)
    #Clicking "Email"
    xpath("//input[@placeholder='Subject']", True, "Here is the Screenshot!")
    #Clicking "Message Body"
    xpath("//div[@aria-label='Message Body']", True, "Thanks for using the program! Application will continue to monitor!")
    #Clicking the Paper Clip
    xpath("//tbody/tr/td/div/div[@aria-label='Attach files']/div/div/div[1]", True)
    #Navigating the Explorer.exe
    screenshot_pyautogui()
    #Clicking "Send"
    time.sleep(10)
    xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
    time.sleep(2)
    driver.close()
    time.sleep(user_time)
    application()

def NO_email_Screenshot(): # This needs a time.sleep function then still continues to monitor the program
    """ This function is for when the user replies with 'NO'. It send another email 
        asking if they want to open it! """
    initialization()
    #Navigating the Gmail Website
    #Clicking "Email or Phone" section + inputting keys(email) for it
    xpath("//input[@id='identifierId']", True, email)
    #Clicking "Next" 
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Password" section + inputting keys for it
    xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
    #Clicking "Next"
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Compose"
    xpath("//div[@role='button'][normalize-space()='Compose']", True)
    #Clicking "To"
    xpath("//input[@role='combobox']", True, to)
    #Clicking "Email"
    xpath("//input[@placeholder='Subject']", True, "Okay! Application will now continue to monitor!")
    #Clicking "Message Body"
    xpath("//div[@aria-label='Message Body']", True, "Thanks for using the program!")
    #Clicking "Send"
    xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
    time.sleep(2)
    driver.close()
    time.sleep(user_time)
    application()

def screenshot():
    """ This function screenshots and saves the file. """
    try:
        time.sleep(3)
        myScreenshot = pyautogui.screenshot()
        myScreenshot.save(f"{path_to_screenshot}\\{name_of_screenshot}")

        print("Screenshot Saved!")
        time.sleep(5)
        Yes_Email_Screeshot()
    
    except:
        error_email()

def application_start():
    """ This function navigates through explorer to find and start the selected application. """
    try:
        os.startfile("Explorer.exe")
        time.sleep(1)
        pyautogui.keyDown("ctrl")
        time.sleep(1)
        pyautogui.press("l")
        time.sleep(1)
        pyautogui.press("ctrl")
        time.sleep(1)
        pyautogui.write(path_to_app)
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.keyDown("ctrl")
        time.sleep(1)
        pyautogui.press("f")
        time.sleep(1)
        pyautogui.keyUp("ctrl")
        time.sleep(1)
        pyautogui.write(file_name)
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(5)
        pyautogui.press("down")
        time.sleep(1)
        #Moving the scroller to the top
        for i in range(100): #Scrolls all the way to the top of the page
            pyautogui.press("up")
        time.sleep(10)
        user_text = False
        while user_text == False:
            pyautogui.press("F2")
            time.sleep(5)
            pyautogui.keyDown("ctrl")
            time.sleep(1)
            pyautogui.press("a")
            time.sleep(1)
            pyautogui.keyUp("ctrl")
            time.sleep(1)
            pyautogui.keyDown("ctrl")
            time.sleep(1)
            pyautogui.press("c")
            time.sleep(1)
            pyautogui.keyUp("ctrl")
            time.sleep(1)
            pyautogui.press("right")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(1)
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            print(data)
            if data != file_name:
                pyautogui.press("down")
                time.sleep(1)
            elif data == file_name:
                    pyautogui.press("enter", presses=2)
                    time.sleep(25)
                    #Turning the file_name into a variable w/o extension
                    notepadWindow = gw.getWindowsWithTitle(taskbar_name)[0]
                    boolean_statement = notepadWindow.isMaximized
                    print(boolean_statement)
                    if boolean_statement == False:
                        notepadWindow.maximize()
                        user_text = True
                    elif boolean_statement == True:
                        user_text = True

        screenshot()
    except:
        error_email()

def not_app_start():
    """ This function emails the user saying the program will not open! """
    initialization()
    #Navigating the Gmail Website
    #Clicking "Email or Phone" section + inputting keys(email) for it
    xpath("//input[@id='identifierId']", True, email)
    #Clicking "Next" 
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Password" section + inputting keys for it
    xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
    #Clicking "Next"
    xpath("//span[normalize-space()='Next']", True)
    #Clicking "Compose"
    xpath("//div[@role='button'][normalize-space()='Compose']", True)
    #Clicking "To"
    xpath("//input[@role='combobox']", True, to)
    #Clicking "Email"
    xpath("//input[@placeholder='Subject']", True, "Okay! program will now close!")
    #Clicking "Message Body"
    xpath("//div[@aria-label='Message Body']", True, "Thanks for using the program!")
    #Clicking "Send"
    xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
    time.sleep(2)
    driver.close()
    print("Exiting")

#4
#Taking note of what the user said
def user_reply():
    """ This function is seeing the users reply! """
    #Seeing if user says Yes or No
    try:
        test = driver.find_element(By.XPATH, "/html/body/div[7]/div[3]/div/div[2]/div[2]/div/div/div/div[2]/div/div[1]/div/div[2]/div/table/tr/td/div[2]/div[2]/div/div[3]/div[2]/div/div/div/div/div[1]/div[2]/div[3]/div[3]/div/div[1]").text

    #These if statements are for the Application Running Section
        if test == f"Yes {password_veri}" or test == f"yes {password_veri}":
            print("Yes, It is working Properly")
            driver.close()
            screenshot()
        elif test == f"No {password_veri}" or test == f"no {password_veri}":
            print("No, It is working properly")
            driver.close()
            NO_email_Screenshot()
        elif test == f"Start app {password_veri}" or test == f"start app {password_veri}":
            print("Starting app")
            driver.close()
            application_start()
        elif test == f"Keep app closed {password_veri}" or test == f"Keep app closed {password_veri}":
            not_app_start()
        elif test == f"Stop program {password_veri}" or test == f"stop program {password_veri}":
            not_app_start()
        else: #This else is something wrong with the text. If the format is wrong or password is wrong, this else will run.
            print("Error_Inside")
            driver.close()
            reaffirmation_email()
#Error
    except:
        print("Error_Outside")
        error_email()

#3
def waitting_for_reply(): #Waiting for a reply from user!
    """ This function is waiting for the users reply. This verifies the username of the reply
        and clicks the image. Then it goes on to Stage two, which is reading the text. """
    try:
        xpath(f"(//span[@name='{username}'][normalize-space()='{username}'])[2]", True)
        time.sleep(30)
        user_reply()
    except:
        error_email()
        
#2
#app_running or not functions
def app_running():
    """ This function sends an Application Running email. """
    try: 
        #Navigating the Gmail Website
        #Clicking "Email or Phone" section + inputting keys(email) for it
        xpath("//input[@id='identifierId']", True, email)
        #Clicking "Next" 
        xpath("//span[normalize-space()='Next']", True)
        #Clicking "Password" section + inputting keys for it
        xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
        #Clicking "Next"
        xpath("//span[normalize-space()='Next']", True)
        #Clicking "Compose"
        xpath("//div[@role='button'][normalize-space()='Compose']", True)
        #Clicking "To"
        xpath("//input[@role='combobox']", True, to)
        #Clicking "Email"
        xpath("//input[@placeholder='Subject']", True, subject_ifapprunning)
        #Clicking "Message Body"
        xpath("//div[@aria-label='Message Body']", True, message_body_ifapprunning)
        #Clicking "Send"
        xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
        time.sleep(60)

        #Calling the Waiting_For_Reply Function
        waitting_for_reply()
    except:
        error_email()

def app_not_running():
    """ This function sends an Application NOT Running email. """
    try: 
        #Navigating the Gmail Website
        #Clicking "Email or Phone" section + inputting keys(email) for it
        xpath("//input[@id='identifierId']", True, email)
        #Clicking "Next" 
        xpath("//span[normalize-space()='Next']", True)
        #Clicking "Password" section + inputting keys for it
        xpath('//*[@id="password"]/div[1]/div/div[1]/input', True, password)
        #Clicking "Next"
        xpath("//span[normalize-space()='Next']", True)
        #Clicking "Compose"
        xpath("//div[@role='button'][normalize-space()='Compose']", True)
        #Clicking "To"
        xpath("//input[@role='combobox']", True, to)
        #Clicking "Email"
        xpath("//input[@placeholder='Subject']", True, subject_ifappNOTrunning)
        #Clicking "Message Body"
        xpath("//div[@aria-label='Message Body']", True, message_body_ifappNOTrunning)
        #Clicking "Send"
        xpath("//div[@aria-label='Send ‪(Ctrl-Enter)‬']", True)
        time.sleep(60)

        #Calling the Waiting_For_Reply App NOT Working Function
        waitting_for_reply()
    except:
        error_email()

#1
def application(): #This is the only function that needs to be called. Every other function gets called inside of functions!
    """ This function is the start of it all! This function also checks the application
        then transfers it to the appropriate function. """
    #Application to Monitor
    try:
        pymem.Pymem(f"{application_monitor}")
        print("App_Running")
        initialization()
        app_running()
    except pymem.exception.ProcessNotFound:
        print("Not Running")
        initialization()
        app_not_running()
#Calling the Application Function
print("Starting up the program in 10 seconds. Please leave the application you want to monitor on the screen!")
time.sleep(10)
application()