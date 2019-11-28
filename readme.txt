Python 3.7.5
Setup:
Ensure that Chromedriver.exe is in the C:\Data folder
Outlook:
setup rules to send all register reports to one folder
The folder these emails are sent to must be a top level folder, meaning it is on the same level as your inbox folder
The easiest method is to use the rules wizard and move all emails from the relevent email addresses to the correct folder.
Simply drag the folder in question to the top where it lists your email (first.last@buschgardens.com) and drop, you should see it move to alphabetical order with your other folders such as inbox and sent items


User Instructions:
Launch Greentag.exe as any other .exe, no admin permissions should be required.
A command prompt will open and request the name of the folder your greentag reports are sent to.
The folder must be a top level (same level as inbox) for the script to reliably find it. See Install for more information

Maintenece and Operation:
Greentag.exe connects to your outlook automatically as long as you are signed in through the Windows win32com package.
Upon opening it will request a folder to search in for Greentag emails. (Capitalization is not important for any user inputs)
This folder must be a top level folder (The same level as Inbox, not inside it) in order to find the folder reliably.
The user will then recieve a prompt for a date of the format YYYY-MM-DD. Leaving the entry blank will use the system's current date.

The script will then encode all emails in that folder from the given date as arrays, and iterate through those arrays looking for register names.
It uses the following Regex identifiers to find and sort registers into their correct park and department:
AITc = "^AIT...[0-9][0-9][0-9]" #Adventure Island Culinary
AITm = "^AIT.....[0-9][0-9][0-9]" #Adventure Island Merchandise
BGTc = "^BGT...[0-9][0-9][0-9]" #Busch Gardens Culinary
BGTm = "^BGT.....[0-9][0-9][0-9]" #BuschGardens Merchandise

The "^" symbol denotes that the string begins with the regex identifier that follows
"."s denote any character
[0-9] indicates any digit
The script then iterates through the next 10 nodes to look for either "Online" or "Offline", if neither is found "ERROR!" will be listed as the register's status.
It is important to note that extremely long location names may break this section of code
In this case, find the line reading "while z < len(strings) and z < i + 10:" and change 10. Lower numbers will yield higher performance, but may miss indicators blocked by names.
The results of the scan are used to create Register objects which hold the name, location, and status of each register
The register objects are then sorted into arrays listing offline registers for each park and department.
Once the full email is parsed into these arrays, the script then prints all registers held in the offline registers array
The script then generates a greentag statement for offline registers to be pasted into the greentag online portal.

The final prompt will ask the user if they would like to begin filling forms
if the user says yes, the program will then begin launching chrome windows which will navigate to the Service Now homepage and then to the new incident page
The script then iterates through each type of offline register and uses that information to fill in the incident tickets.
The script is hard coded with names to correspond with each park and department combination to fill in the callerID, and uses the Register datapoints to fill in the remaining information
At current the user is required to cross reference with AD to find the location actual and paste that into the short description on Service Now, Future iterations may link with AD to provide the location inside the report.
Users should check the fields filled in for any errors before submiting.

This script relies on Selenium to launch chrome from Chromedriver.exe. Chromedriver.exe is a lightweight version of Google chrome used for automation.
Chromedriver pulls from the currently installed version of Google Chrome and must be updated from the Selenium website if chrome is updated.
If the script fails when running the form filler, check the version of chrome by opening a chrome window, opening the triple dot menu > help > about chrome.
Document any updates to Chromedriver.exe in this Readme
Current Chromedriver.exe version: 74




