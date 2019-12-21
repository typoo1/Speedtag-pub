import os
import os.path
import sys
import re
import datetime
import selenium
import threading
import logging
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import win32com
import win32com.client
import win32timezone
from win32com.client import Dispatch
import pyad
from pyad import adquery
from selenium.webdriver.common.keys import Keys
import itertools as it

"""
Program by Tye Alexander Gallagher
This program automates the morning greentagging proess for registers by reading in the user's emails and parsing the reports therein.
The program then uses that information to create a string to be pasted into the greentag portal and has the option of filling out incident reports to match those reports.
"""

Cities = ["TMP", "ORL", "SDO", "SAT", "LAG", "WIL"]
global city

CustomerNum = 4

"""
Default customer names
"""
TMPCus = ["Samuel Goldstein", "Tyrone Robinson", "Terrence Lattimore", "Sidra Davis", "blank", "blank"]
ORLCus = ["Samuel Goldstein", "Tyrone Robinson", "Terrence Lattimore", "Sidra Davis", "blank", "blank"]
SDOCus = ["Samuel Goldstein", "Tyrone Robinson", "Terrence Lattimore", "Sidra Davis", "blank", "blank"]
SATCus = ["John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins", "John Blevins"]
LAGCus = ["Samuel Goldstein", "Tyrone Robinson", "Terrence Lattimore", "Sidra Davis", "blank", "blank"]
WILCus = ["Samuel Goldstein", "Tyrone Robinson", "Terrence Lattimore", "Sidra Davis", "blank", "blank"]

Cus = [] #Final customer list, to be loaded during Form Filler

q = pyad.adquery.ADQuery()

"""
Global variables to hold  form filler items
"""
global P1CusC #Park 1 Culinary Customer
global P2CusC #Park 2 Culinary Customer
global P1CusM #Park 1 Merch Customer
global P1CusC #Park 2 Merch Customer
global AGroup #Assignment group
global P1ADOU #Holds the OU for park 1

"""
Grabs the path the program is currently in and adds a "\", used to find other necissary files
"""
path = os.path.dirname(os.path.abspath('Main.py'))
path = path + "\\"

"""
Arrays for holding various register objects
"""
offlineReg = []
probReg = []
offlineRemQueue = []
xstoreRegaR = []
xstoreRegbR = []
xstoreRegbM = []
xstoreRegaM = []
culinaryRega = []
culinaryRegb = []
MPRRega = []
MPRRegb = []


"""
Sets the path for the chromedriver and configures the options so that chrome does not close once the page is loaded.
"""
driverP = (path + r"chromedriver.exe") 
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument('log-level=3')
chrome_options.add_argument('disable-infobars')
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')

offlineC = 0 #Count for offline registers

registers = [] #Array that holds all registers

f = open("testfile.txt", "w+") #test file for debugging purposes

"""
Attaches to client outlook to grab emails
"""
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts

global tFolder #Target search folder

"""
END OF VARIABLE SETUP
"""

"""
Sets the customer values according sure the config file.
"""
def setCus():
    global P1CusC
    global P2CusC
    global P1CusM
    global P2CusM
    global P1cusMPR
    global P2cusMPR
    P1CusC = Cus[0]
    P2CusC = Cus[1]
    P1CusM = Cus[2]
    P2CusM = Cus[3]
    P1cusMPR = Cus[4]
    P2cusMPR = Cus[5]

"""
Takes in a string as an argument and uses it to determine the city the user is configuring their client for.
This then assigns all Park code values, full names, assignments groups, and creates the Regex strings for find register names in the emails.
"""
def setPark(City):
    global Park1
    global Park2
    global Park1c
    global Park1m
    global Park2c
    global Park2m
    global Park1MPR
    global Park2MPR
    global Park1QQ
    global Park2QQ
    global Park1Full
    global Park2Full
    global AGroup
    global city
    global P1ADOU
    global P2ADOU
    city = City
    
    
    print("Setting city to " + city)
    if(city == "TMP"): #Tampa
            Park1 = "BGT"
            Park2 = "AIT"
            Park1Full = "Busch Gardens Tampa"
            Park2Full = "Adventure Island Tampa"
            AGroup = "BGT-IT"
            P1ADOU = "_BGT"
                
    elif(city == "ORL"): #Orlando
            Park1 = "SWF"
            Park2 = "APO"
            Park1Full = "SeaWorld Florida"
            Park2Full = "Aquatica Park Florida"
            AGroup = "SWF-IT"
            P1ADOU = "_SWF"
            
    elif(city == "SDO"): #San Diego
            Park1 = "SWC"
            Park2 = "APC"
            Park1Full = "SeaWorld California"
            Park2Full = "Aquatica Park California"
            AGroup = "SWC-IT"
            P1ADOU = "_SWC"
            
    elif(city == "SAT"): #San Antonio
            Park1 = "SWT"
            Park2 = "APT"
            Park1Full = "SeaWorld Texas"
            Park2Full = "Aquatica Park Texas"
            AGroup = "SWT-IT"
            P1ADOU = "_SWT"
            
    elif(city == "LAG"): #Langhorn
            Park1 = "SPL"
            Park2 = ""
            Park1Full = "Sesame Place Langhorn"
            Park2Full = ""
            AGroup = "SPL-IT"
            P1ADOU = "_SPL"

    elif(city == "WIL"): #Williamsburg
            Park1 = "BGW"
            Park2 = "WCW"
            Park1Full = "Busch Gardens Williamsburg"
            Park2Full = "Water Country USA"
            AGroup = "BGW-IT"
            P1ADOU = "_BGW"

    Park1c = "^" + Park1 + "...[0-9][0-9][0-9]"
    Park1m = "^" + Park1 + ".....[0-9][0-9][0-9]"
    Park2c = "^" + Park2 + "...[0-9][0-9][0-9]"
    Park2m = "^" + Park2 + ".....[0-9][0-9][0-9]"
    Park1MPR = "^" + Park1 + ".*MPR[0-9][0-9][0-9]"
    Park2MPR = "^" + Park2 + ".*MPR[0-9][0-9][0-9]"
    Park1QQ = "^" + Park1 + ".*QQ.*[0-9][0-9][0-9]"
    Park2QQ = "^" + Park2 + ".*QQ.*[0-9][0-9][0-9]"

"""

"""
def printCulGreB(reg):
    if reg:
        s = Park1 + "RCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

"""

"""
def printCulGreA(reg):
    if reg:
        s = Park2 + "RCP"
        for regX in reg:
            s = s + regX.name[-3:] + ", "
        s = s[:-2]
        s = s + " offline."
        return s

"""

"""
def xStoreMB(reg):
    if reg:
        s = ""
        for reg in reg:
            s = s + reg.name[-3:] + ", "
        s = s[:-2]
        return s
"""
Takes a list of registers in and creates a list of all registers
"""
def xStoreRB(reg):
    s = ""
    for reg in reg:
        s = s + reg.name[-3:] + ", "
    s = s[:-2]
    return s


"""

"""
def printXStoreA(regR, regM):
    if reg:
        s = ""
        if (regR and regM):
            s = Park2 + "RMPOS" + xStoreRB(regR) + " offline; " + Park2 + "MMPOS" + xStoreMB(regM)
        elif (regR):
            s = Park2 + "RMPOS" + xStoreRB(regR) + " offline."
        elif (regM):
            s = Park2 + "MMPOS" + xStoreMB(regM) + " offline."
        return s

"""

"""
def printXStoreB(regR, regM):
    s= ""
    if(regR and regM):
        s = Park1 + "RMPOS" + xStoreRB(regR) + " offline; " + Park1 + "MMPOS" + xStoreMB(regM) + " offline."
    elif(regR):
        s = Park1 + "RMPOS" + xStoreRB(regR) + " offline."
    elif(regM):
        s = Park1 + "MMPOS" + xStoreMB(regM) + " offline."
    return s

def printMPRGreA(reg):
    s = ""
    for reg in reg:
        s = s + reg.name + ", "
    s = s[:-2]
    s = s + " offline."
    return s

def printMPRGreB(reg):
    s = ""
    for reg in reg:
        s = s + reg.name + ", "
    s = s[:-2]
    s = s + " offline."
    return s

"""
Determines wich print statements to run and prints the results for the user. This result is used to create the greentag statement to copy-paste into the greentag portal.
"""
def printReg():
    if(culinaryRegb):
        print(printCulGreB(culinaryRegb))
    if(culinaryRega):
        print(printCulGreA(culinaryRega))
    if(xstoreRegbM or xstoreRegbR):
        print(printXStoreB(xstoreRegbR, xstoreRegbM))
    if(xstoreRegaM or xstoreRegaR):
        print(printXStoreB(xstoreRegaR, xstoreRegaM))
    if(MPRRegb):
        print(printMPRGreB(MPRRegb))
    if(MPRRega):
        print(printMPRGreA(MPRRega))
"""
Manages the entire process of opening the browser, navigating to the new incident page, and creating the new incidents
"""
def fillForms():
    threads = []
    if(len(offlineReg) - len(probReg) <= 10):
        for reg in it.chain(offlineReg, probReg):
            x = threading.Thread(target=Forms, args=([reg]))
            x.start()
            threads.append(x)
            time.sleep(1)
        for thread in threads:
            thread.join()
            print(str(thread.name) + " done")
    else:
        for reg in probReg:
            offlineReg.append(reg)
        input("the program can be somewhat unstable when opening many chrome windows at once. \nPress enter to open the first 10 chrome windows")
        i = 0
        while i < len(offlineReg):
            x = threading.Thread(target=Forms, args=([offlineReg[i]]))
            x.start()
            threads.append(x)
            time.sleep(1)

            if((i+1) % 10 == 0):
                for thread in threads:
                    thread.join()
                    print(str(thread.name) + " done")
                    threads.remove(thread)
                input("Please submit and close all open chrome windows, then press enter to continue")
                
            i += 1
        
        

def Forms(reg):
    driver = webdriver.Chrome(driverP, options=chrome_options) #Pass chrome driver the options previous configured
    driver.set_page_load_timeout(300) #Set maximum time that webpage is left to load
    driver.implicitly_wait(60)
    driver.get("https://sea.service-now.com/home.do") #It is necissary to navigate to the Service now homepage to prevent a series of redirects which breaks the remainders of the script
    driver.get("https://sea.service-now.com/incident.do") #Navigate to the New incident form.
    driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
    time.sleep(1)
    CID = driver.find_element_by_name("sys_display.incident.caller_id")  # caller
    REGNAME = driver.find_element_by_name("sys_display.incident.cmdb_ci")  # register name
    CONTYPE = driver.find_element_by_name("incident.contact_type")  # Contact type
    CONTYPE.send_keys("Email") #Static, does not change
    SHORTDES = driver.find_element_by_name("incident.short_description")  # Short description
    GROUP = driver.find_element_by_name("sys_display.incident.assignment_group") #Assignment group
    if(reg.park == Park2Full): #Check if park2
        if(re.search(Park2c, reg.name)):
            CID.send_keys(P2CusC)
        elif(re.search(Park2m, reg.name)):
            CID.send_keys(P2CusM)
        elif(re.search(Park2MPR, reg.name) or re.search(Park2QQ, reg.name)):
            CID.send_keys(P2cusMPR)
        else:
            CID.send_keys("!!!!ERROR!!!!")
    elif (reg.park == Park1Full): #check if park1
        if (re.search(Park1c, reg.name)):
            CID.send_keys(P1CusC)
        elif (re.search(Park1m, reg.name)):
            CID.send_keys(P1CusM)
        elif(re.search(Park1MPR, reg.name) or re.search(Park1QQ, reg.name)):
            CID.send_keys(P1cusMPR)
        else:
            CID.send_keys("!!!!ERROR!!!!")
    #CID.send_keys("\ue003")  # down arrow ##Depricated, caused issues with omitting the final character initally and no longer needed
    CID.send_keys("\ue004")  # tab
    REGNAME.send_keys(reg.name)
    REGNAME.send_keys("\ue004")  # tab
    if(reg.status.lower() == "offline"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " offline on morning report.")
    if(reg.status == "HDD problem"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " is low on HDD space.")
    if(reg.status == "Repl problem"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " is experiencing Replication error.")
    if(reg.status == "Close Failure"):
        SHORTDES.send_keys(reg.name + " " + reg.loc + " experienced a store close failure.")
    SHORTDES.send_keys("\ue004")  # tab
    REGNAME.send_keys(Keys.BACKSPACE)
    time.sleep(1)
    REGNAME.send_keys(reg.name[-1:])
    time.sleep(1)
##    REGNAME.send_keys(Keys.DOWN)
##    REGNAME.send_keys("\ue004")  # tab
    GROUP.send_keys((Keys.CONTROL + "a"))
    time.sleep(1)
    GROUP.send_keys(Park1 + "-IT") #should be filling automatically based on CI
    GROUP.send_keys("\ue004")
    #driver.implicitly_wait(1) #Wait statement, no longer needed.
"""
Defines the structure of the Register object used to store register information
"""
class Register:
    
    """
    Default Constructor
    """
    def __init__(self):
        self.name = ""
        self.park = ""
        self.status = ""
        self.loc = ""
        self.HDD = 100.0
        
##    """
##    Constructor with two arguments for the name and reported status of the register
##    """
##    def __init__(self, name, status):
##        self.name = name
##        x = ""
##        if(re.search("^" + Park2, name)):
##            self.park = Park2Full
##        else:
##            if(re.search("^" + Park1, name)):
##                self.park = Park1Full
##            else:
##                self.park = "ERROR"
##        if(status == "off"):
##                status = "offline"
##        if(status == "online" or status == "offline" or status == "off"):
##            self.status = status
##            self.loc= ""
##            if(status == "offline"):
##                self.setLoc()
##        else:
##                self.status = "ERROR"
##                print("REGISTER STATUS NOT FOUND, PLEASE REVIEW " + self.name)
##        self.HDD = 100.0

    """
    Constructor with added HDD argument. This is specific to the Culinary reports for detecting HDD faults.
    """
    def __init__(self, name, status, HDDVal):
        self.name = name
        x = ""
        if(re.search("^" + Park2, name)):
            self.park = Park2Full
        else:
            if(re.search("^" + Park1, name)):
                self.park = Park1Full
            else:
                self.park = "ERROR"
        if(status == "off"):
                status = "offline"
        if(status == "online" or status == "offline" or status == "HDD problem" or status == "Repl problem" or  status == "Close Failure"):
            self.status = status
            self.loc= ""
            if(status != "online"):
                self.setLoc()
        else:
                self.status = "ERROR"
                print("REGISTER STATUS NOT FOUND, PLEASE REVIEW " + self.name)
        self.HDD = HDDVal
                
    """
    Method that defines how to print information stored in the Register object
    """
    def printReg(self):
        if(self.status == "offline"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is " + str(self.status))
        if(self.status == "HDD problem"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is low on HDD space.")
        if(self.status == "Repl problem"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " is experiencing a replication problem.")
        if(self.status == "Close Failure"):
            print("Register " + self.name + " at " + self.park + ": " + self.loc + " experienced a store close failure.")
            

    def setLoc(self):
        if(re.search("MMPOS", self.name) or re.search("RMPOS", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=Xstore,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        elif(re.search("CP", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=Culinary_POS,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        elif(re.search("MPR", self.name) or re.search("QQ", self.name)):
            q.execute_query(
                attributes = ["CN", "description"],
                where_clause = "CN = '" + self.name + "'",
                base_dn = "OU=MPR,OU=_Special Needs - Advertise & Install - No Reboot,OU=Computers,OU=" + P1ADOU + ",DC=nam,DC=int,DC=local")
        for row in q.get_results():
            self.loc = str(row["description"])[2:-3]

"""
Responsible for parsing emails
Takes a folder item as input to scan through
This method grabs all the messages from a folder without reviewing their contents, sorts them by their SentOn value, and iterates through them until a SentOn value doesn't match the current system date before breaking.
As the method iterates, it checks the sender to ensure that it is from a known source of greentag emails to avoid processing info from irrelevent emails.
"""
def emailleri_al(folder):
    messages = folder.Items
    a=len(messages)
    print("Parsing the following emails...\n")
    messages = sorted(messages, key=lambda messages: messages.SentOn)
    if a>0:
        for message2 in reversed(messages):
            try:
                sender = message2.SenderEmailAddress.lower()
                sdate = str(message2.SentOn)
                if sender != "":
                    if sender == "seap2018@seaworld.com" or "xstorereport@seaworld.com" or "swt.ithelpdesk@SeaWorld.com": #TODO find correct email addresses"
                        if sdate[:-15] == str(tarDate):
                            print(message2.Subject)
                            print("***********************", file =f)
                            print(message2.Subject, file=f)
                            print("***********************", file=f)
                            print(str(message2.SentOn)[:-15], file=f)
                            #manipulate the string
                            output = message2.Body
                            output = ' '.join(output.split())
                            print(output, file=f)
                            #Create register objects
                            strings = []
                            strings = output.split()
                            #print(strings)
                            #print(len(strings))
                            i = 0
                            
                            while i < len(strings):
                                if(re.search(Park2c, strings[i]) or re.search(Park2m, strings[i]) or re.search(Park1c, strings[i]) or re.search(Park1MPR, strings[i]) or re.search(Park2MPR, strings[i])
                                        or re.search(Park1m, strings[i]) or re.search(Park1QQ, strings[i]) or re.search(Park2QQ, strings[i])):
                                    z = i
                                    statSet = False
                                    while z < len(strings) and z < i + 20:
                                        if(strings[z].lower() == "online" or strings[z].lower() == "offline" or strings[z].lower() == "off"):
                                            if(statSet): #should only trigger on xstore registers
                                                if(strings[z-1].lower() == "failed"):
                                                    registers.append(Register(strings[i], "Close Failure", 101))
                                                registers.append(Register(strings[i], status, 101))
                                                break #This is used to prevent lines overlapping on Xstore reports specifically
                                            status = strings[z].lower()
                                            statSet = True
                                            #break
                                        if(strings[z] == "%"): #Detect if register is from FreedomPay report dynamically
                                           if(float(strings[z-1]) < 20):
                                                registers.append(Register(strings[i], "HDD problem", float(strings[z-1])))
                                                
                                           if(not((strings[z-5] == "2" or strings[z-6] == "2") or (strings[z-5] == "1" or strings[z-6] == "1"))): #if repl status != 1 or 2
                                                registers.append(Register(strings[i], "Repl problem", 100))
                                                
                                           registers.append(Register(strings[i], status, float(strings[z-1])))
                                           break
                                        z += 1
                                    # print("**************" + registers[i].name)
                                    # regNum = regNum + 1
                                i = i + 1
                        else:
                            print()
                            break
                    print(sender, file=f)
                try:
                    message2.Save
                    message2.Close(0)
                except:
                    pass
            except:
                print(sys.exc_info())
                pass
"""
Uses the win32com package to check the version of Chrome currently installed and throws an error if it's incompatible with the current chromedriver.exe installed
"""
def getVer():
    parser = Dispatch("Scripting.FileSystemObject")
    version = parser.GetFileVersion(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    if not(version[:2] == "74"):
        raise ValueError("You are using an invalid version of Chrome, please update to version 74")
                
"""
Checks to see if a config file exiests yet, and if not goes through several prompts with the user to create the appropriate config for them.
The parser for the config file assumes that the title of each config item is only a single word long.
"""
def getConfig():
    global tFolder
    global Cus
    if not (os.path.exists(path + "config.txt")):
        print("Config file not found, performing first time setup")
        config = open("config.txt", "w+")
        x = input("Please enter the name of the folder your register reports are put in ").lower()
        config.write("OutlookFolder: " + x + "\n")
        tFolder = x
        print("the target folder is: " + tFolder)
        x = input("Please enter your city code: TMP, ORL, SDO, SAT, LAG, WIL ").upper()
        config.write("CityCode: " + x + "\n")
        setPark(x)
        i = 0
        if(x == "TMP" or "SAT"):
            if(input("Would you like to use the default customers for your park? y/n").lower() == "y"):
                if(city == "TMP"):
                    Cus = TMPCus
                if(city == "ORL"):
                    Cus = ORLCus
                if(city == "SDO"):
                    Cus = SDOCus
                if(city == "SAT"):
                    Cus = SATCus
                if(city == "LAG"):
                    Cus = LAGCus
                if(city == "WIL"):
                    Cus = WILCus
            else:
                print("beginning custom customer setup...")
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Culinary: "))
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Culinary: "))
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Merchandise: "))
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Merchandise: "))
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " MPR (Leave blank if you don't recieve this report): "))
                Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " MPR (Leave blank if you don't recieve this report): "))
        else:
            print("Your parks do not currently have default customer values. Please email Tye.Gallagher@BuschGardens.com with your prefered values to have them implimented")
            print("beginning custom customer setup...")
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Culinary: "))
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Culinary: "))
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " Merchandise: "))
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " Merchandise: "))
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park1Full + " MPR (Leave blank if you don't recieve this report): "))
            Cus.append(input("Enter the name as it appears in Service Now of the person responsible for " + Park2Full + " MPR (Leave blank if you don't recieve this report): "))
        for p in Cus:
            config.write("Customer" + str(i) + ": " + p +"\n")
            i += 1
        setCus()
        print("config created")
        config.close()
    else:
        try:
            config = open("config.txt", "r")
            y = []
            x = config.readlines()
            l = len(x)
            i = 0
            for item in x:
                y.append(parseItem(item))#load y with config variables
                i+=1
            tFolder = y[0] #assign the first item in the config as the target folder
            setPark(y[1]) #asign the second item as the city
            i = 2 
            while i < len(y):
                Cus.append(y[i]) #assign the remaining items as the customers to be used
                i += 1
            setCus()
            print("config loaded")
        except:
            raise ValueError("Something went wrong with your config file, either edit it or delete it to have the script replace it.")
        config.close()
"""
Method used to parse the individual lines of the config file. It takes every part of a given line, removes the first item (the item's title) and returns the remaining items seperated by spaces as a single string
"""
def parseItem(item):
    y = item.split()
    l = len(y)
    i = 1
    result = ""
    while i < l:
        result = result + y[i] + " "
        i += 1
    result = result[:-1]
    return result
"""
Takes the in an account as an argument, then scans through all the account's email folders and their subfolders to find the one matching the target folder.
The method then calls the emailleri_al method on the target folder to scan for the desired emails.
"""
def getEmails(Account):
    for account in accounts:
        global inbox
        global tarDate
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        print("****Account Name*********************", file=f)
        print(account.DisplayName, file=f)
        print("From " + account.DisplayName +"\n")
        print("*************************************", file=f)
        folders = inbox.Folders
        tarDate = datetime.date.today()
        print("Processing mail from the " + tFolder + " folder from " + str(tarDate) + "\n")

        for folder in folders:
            print("*****Folder Name*****************", file=f)
            print(folder, file=f)
            print("*********************************", file=f)
            if(str(folder).lower() == tFolder.lower()):
                emailleri_al(folder)
            a = len(folder.folders)
            if a>0 :
                global z
                z = outlook.Folders(account.DeliveryStore.DisplayName).Folders(folder.name)
                x = z.Folders
                for y in x:
                    print("*****Folder Name********************", file=f)
                    print("..."+y.name, file=f)
                    print("************************************", file=f)
                    if(str(y).lower() == tFolder.lower()):
                        emailleri_al(y)

def Ping(name, txt):
    print(name, file=txt)
    print(name + " ping status " + str(os.system("Echo off & ping -n 1 " + name + " > NUL")))


def PrintOffline():
    global offlineReg
    offTxt = open("Offline.txt", "w+")
    print("Printing ping results...")
    threads = []
    for Reg in offlineReg:
        x = threading.Thread(target = Ping, args=(Reg.name, offTxt))
        x.start()
        threads.append(x)
    for thread in threads:
        thread.join()


"""
END OF DEFS
"""

print("Speedtag 1.91V")
print("A Program by Tye A. Gallagher")
print()
getVer()
getConfig()
getEmails(accounts)

for reg in registers:
    if(reg.status.lower() == "offline" or reg.status.upper() == "ERROR"):
        reg.printReg()
        offlineC += 1
        offlineReg.append(reg)
    if(reg.status == "HDD problem" or reg.status == "Repl problem" or reg.status == "Close Failure"):
        reg.printReg()
        offlineC += 1
        probReg.append(reg)
print()
print("There are " + str(len(registers)) + " registers reported on and " + str(offlineC) + " registers that need attention")

for reg in offlineReg:
    if(re.search("^" + Park1 + "RMPOS", reg.name)):
        xstoreRegbR.append(reg)
    if(re.search("^" + Park1 + "MMPOS", reg.name)):
        xstoreRegbM.append(reg)
    if(re.search("^" + Park2 + "RMPOS", reg.name)):
        xstoreRega.append(reg)
    if(re.search("^" + Park2 + "MMPOS", reg.name)):
        xstoreRegaM.append(reg)
    if(re.search("^" + Park1 + "RCP", reg.name) or (re.search("^" + Park1 + "CP", reg.name) or (re.search("^" + Park1 + "MCP", reg.name)))):
        culinaryRegb.append(reg)
    if(re.search("^" + Park2 + "RCP", reg.name) or (re.search("^" + Park2 + "CP", reg.name) or (re.search("^" + Park2 + "MCP", reg.name)))):
        culinaryRega.append(reg)
    if(re.search("^" + Park1 + ".*MPR", reg.name) or re.search("^" + Park1 + ".*QQ", reg.name)):
       MPRRegb.append(reg)
    if(re.search("^" + Park2 + ".*MPR", reg.name) or re.search("^" + Park2 + ".*QQ", reg.name)):
       MPRRega.append(reg)
    if(reg.loc[7:].lower() == "offline"):
        print("removing " + reg.name + " as offline")
        offlineRemQueue.append(reg)

for reg in offlineRemQueue:
    offlineReg.remove(reg)

print()
printReg()
#PrintOffline() #TODO make the print statements come out nicer

if(input("\nProceed with form filler? Y/N \n").lower() == "y"):
    fillForms()

"""
END OF PROGRAM
"""
