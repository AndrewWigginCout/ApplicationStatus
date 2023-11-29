""" This is the only file you should be editing! Every continuous dash is a section! 
    This file was made with the user in mind. If you have no Python experience, there's
    no need to go digging through the code. You're free to look at it though! 
    If you're confused or overwhelmed, please open the Selenium Project Notes file
    and read the 'user_information' file section. (Last section of the notes) """

#NOTE: YOU ONLY NEED TO EDIT THE EMPTY QUOTES WITH INFORMATION! The green text is short information. Again, for the long explanation,
#      Please open the Selenium Project Notes file and edit this file accordingly! 


#Application Section
application_monitor = "" #Application to monitor
file_name = "" #This is if the file name is different than the process name. If they're the same you still NEED to add both
taskbar_name = "" #This is so getWindowsWithTitle can run perfectly!

#Credentials Section
email = "" #Email

password = "" #Password

#Sending Email Section
to = "" #This is a variable for who to send the email to

subject_ifapprunning = "Application Status: Working!" #Subject if the application is running!

message_body_ifapprunning = "Application Status: Working! Do you want to send a screenshot? Reply back with yes or no with the password you assigned. The reply should look like this, yes password. To stop program, type in Stop program password" #Body if the application is running

subject_ifappNOTrunning = "Application Status: Not Working.." #Subject if the app IS running

message_body_ifappNOTrunning = "Application Status: Not Working.. Do you want to run the app? If yes, reply back with 'start app w/ password' The reply should look like this, Start app password'. \
                                \
                                \
If not, reply back with 'Keep app closed password'(The program will come to an end). To stop program, type in Stop program password" #Body if the app is NOT running


#Verification Section
username = "" #Username verification

password_veri = "" #Password verification

#Path to application
path_to_app = "" #This is a bit more complitcated. Please see Selenium Notes!

#Path to save screenshot
name_of_screenshot = "" #This is a custom name
path_to_screenshot = "" #This is the location of it

#User_Inputted time
user_time = 30 #This is the time variable in seconds. This lets the program know how long between emails to send another one. (calculated in seconds. 3600 seconds = hour) (This has no quotes since it's an integar!)

g={}
for variable in dir():
  if not variable.startswith('__') and variable!='g': 
    g[variable]=eval(variable)
print(g)
import json
with open('config.json','w') as fh:
  fh.write(json.dumps(g,indent=2))