Python 3.7.5
Speedtag2.0

Setup:
Ensure that Chromedriver.exe is in the same folder as SpeedTag2.0.exe
Outlook:
Setup rules to send all register reports to one folder
The easiest method is to use the rules wizard and move all emails from the relevent email addresses to the correct folder.


User Instructions:
Launch SpeedTag1.8.exe as any other .exe, no admin permissions should be required.
A command prompt will open. Depending on if you have a config file set up yet or not, said file will either be loaded, or you'll be prompted to create the file through the script.
If an error occurs around the config file, it is recommended to delete the file and try to set it up again.
If the error persists, email Tye.Gallagher@buschgardens.com

Maintenece and Operation:
Speedtag1.8 connects to your outlook automatically, as long as you are signed in, through the Windows win32com package.
Upon opening the first time it will request a folder to search in for Greentag emails. (Capitalization is not important for this or any other user inputs)
The program will only look at emails sent on the current sytem date and explicitly state which emails it will open.

During the first use, or if no config file is found, the script will ask a series of questions that will be used by the script.
This includes the city the park is located in, the folder the script will look in for the reports, and the names of customers the user would like incidents to be called in from.
It is possible to modify the config file manually. The first value (IE CityCode) should not be changed, only values after the space. If any errors occur following modification, it is recommended taht you delete the config and start with a fresh setup.

For reference the city codes are as follows:
TMP = Tampa
ORL = Orlando
SDO = San Diego
SAT = San Antonio
LAG = Langhorn
WIL = Williamsburg

The script will go through the target folder and pull any emails from the current date sent from "seap2018@seaworld.com" or "xstorereport@seaworld.com" or "swt.ithelpdesk@SeaWorld.com".
These emails are then parsed into an array of strings which the script iterates through looking for registers and their status
It uses the following Regex identifiers to find and sort registers into their correct park and department:
    Park1c = "^" + Park1 + ".*CP[0-9]*"
    Park1m = "^" + Park1 + ".*POS[0-9]*"
    Park2c = "^" + Park2 + ".*CP[0-9]*"
    Park2m = "^" + Park2 + ".*POS[0-9]*"
    Park1MPR = "^" + Park1 + ".*MPR[0-9]*"
    Park2MPR = "^" + Park2 + ".*MPR[0-9]*"
    Park1QQ = "^" + Park1 + ".*QQ.*[0-9]*"
    Park2QQ = "^" + Park2 + ".*QQ.*[0-9]*"

where Park1 and Park2 refer to the 2 parks in association with the city given at setup (For Langhorn Park2 is an empty string)
The "^" symbol denotes that the string begins with the regex identifier that follows
"."s denote any character
the * indicates that any number of the previous character or symbol can be between the previous and next symbols
[0-9] indicates any digit

The script then iterates through the next 20 nodes to look for either "Online" or "Offline", if neither is found "ERROR!" will be listed as the register's status.
It is important to note that extremely long location names may break this section of code
In this case, find the line reading "while z < len(strings) and z < i + 20:" and change 20. Lower numbers will yield higher performance, but may miss indicators blocked by names.
The Script also looks for the last character in row (the character is consistent across each report) and uses this to find the repl status, store close failures, and the HDD freespace %
The results of the scan are used to create Register objects which hold the name, location, and status of each register.

The register objects are then sorted into arrays listing offline registers for each park and department and a seperate array for registers with different problems.
The script currently detects abnormal replication statuses and Low disk space on culinary registers as additional problems.
Once the full email is parsed into these arrays, the script then prints all registers held in the offline and problem registers arrays.
The script then generates a greentag statement for offline registers to be pasted into the greentag online portal.

The script will then query AD to find the AD description for objects matching the name of the registeres listed as offline or problematic.
It uses this information to set a location for all the offline registers only to improve performance and reduce stress on AD.
The program will then remove any locations from the offline report that begin with "offline" to save on processing incidents that aren't required or are already open.

The final prompt will ask the user if they would like to begin filling forms.
if the user says yes, the program will then begin launching chrome windows which will navigate to the Service Now homepage and then to the new incident page
The script then iterates through each type of offline register and uses that information to fill in the incident tickets.
The script pulls from the config file for names that correspond with each park and department combination to fill in the callerID, and uses the Register datapoints to fill in the remaining information
Users should check the fields filled in for any errors before submiting.

This script relies on Selenium to launch chrome from Chromedriver.exe. Chromedriver.exe is a lightweight version of Google chrome used for automation.
Chromedriver pulls from the currently installed version of Google Chrome and must be updated from the Selenium website if chrome is updated.
If the script fails when running the form filler, check the version of chrome by opening a chrome window, opening the triple dot menu > help > about chrome.
Document any updates to Chromedriver.exe in this Readme

Current Chromedriver.exe version: 74

If the script closes in the middle of an operation, it is likely that a catostrophic error has occured. In order to view the error, open a command prompt and navigate to the folder the script is installed in.
Run the script from the command prompt. Things should run as normal, if an error occurs the command prompt will, however, stay open allowing you to read the error code. If you require further assistance, email Tye.Gallagher@BuschGardens.com

"# greentag" 