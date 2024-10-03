from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread
import time, os
from subprocess import CREATE_NO_WINDOW
UserEmail = "eplintern.asher@gmail.com"
Password = "abc123"
ChromeOptions = webdriver.ChromeOptions()

def GenNewPW(Creds:dict,PW=None):
    import string, secrets
    specialChar = "-_|/.;][+?"
    alphabet = string.ascii_letters + string.digits + specialChar
    NumPW = len(Creds["Emails"])
    NewDict = {"Instance Name":Creds['Instance Name'],"Emails": Creds["Emails"]}
    Creds.pop("Instance Name")
    Creds.pop("Emails")
    for line in Creds:
        Passwords = []
        while len(Passwords) < NumPW:
            while True:
                password = ''.join(secrets.choice(alphabet) for i in range(15))
                if (any(c.islower() for c in password)
                        and any(c.isupper() for c in password)
                        and sum(c.isdigit() for c in password) >= 3):
                    if password[0] not in specialChar:
                        break
            Passwords.append(password)
        NewDict[line] = Passwords
    return NewDict

def GetCredsFromExcel(File):
    import xlwings as xw
    app = xw.App(visible=False,add_book=False)
    ws = app.books.open(File)
    RResults = ws.sheets[ws.sheets[0]].range("A1:F200").value
    Results = []
    for i in RResults:
        new = []
        if i[0] != None:
            Results.append(i)
    Creds = {}
    for i in Results:
        if i is not None:
            Creds[i[0]] = (i[1:])            
    
    Emails = []
    for i in Creds['Emails']:
        if i is not None:
            Emails.append(i)
    Creds["Emails"] = Emails
    
    ws.close()
    app.quit()
    return Creds

def DictToExcel(dic,OutputPath="/SeleniumPasswordUpdater/Files/",Name=None):
    import datetime,pandas
    if Name == None:
        ExcelName = f"NewPWs-{datetime.datetime.today().strftime('%d-%m-%Y')}.xlsx"
    ExcelFilePath = os.getcwd().replace("\\","/") + OutputPath + ExcelName
    df = (pandas.DataFrame(dic).T).to_excel(ExcelFilePath)
    return ExcelFilePath

def Login(UserEmail, Password, LoginPortal, driver, waitTime = 20):
    wait = WebDriverWait(driver, waitTime)
    #check for SSL web error
    increment = 0
    while increment < 3:
        try:
            #login and navigate to the user management page 
            time.sleep(3)   
            SSLWebError = driver.find_elements(By.ID ,"details-button")
            if len(SSLWebError) > 0:
                SSLWebError[0].click()
                wait.until(EC.element_to_be_clickable((By.ID ,"proceed-link"))).click() 
            
            if len(driver.find_elements(By.XPATH,"//*[text() = 'Odoo Client Error']")) > 0:
                driver.find_element(By.XPATH,"//*[@id='dialog_1']/div/div/div/header/button").click()

            ErrorMessage = driver.find_elements(By.XPATH, "//*[@id='wrapwrap']/main/div/div/div/form/p")
            if len(ErrorMessage) > 0:
                if "Too many login failures, please wait a bit before trying again." in ErrorMessage[0].text:
                    return False
            
            for int in range(3):
                if driver.current_url == LoginPortal:
                    Login = wait.until(EC.element_to_be_clickable((By.NAME, "login")))
                    Login.clear()
                    Login.send_keys(UserEmail)
                    PassWord = wait.until(EC.element_to_be_clickable((By.NAME,"password")))
                    PassWord.clear()
                    PassWord.send_keys(Password)
                    wait.until(EC.element_to_be_clickable((By.XPATH,"//button[text() = 'Log in']"))).click()
                    if len(driver.find_elements(By.XPATH, "//*[@title = 'Home Menu']")) < 1:
                        driver.refresh()
                        time.sleep(3)
                else:
                    driver.refresh()
                    return True
            return False
        except:
            #checks if there is an error page and goes back or reloads if not in login page      
            if len(driver.find_elements(By.NAME,"login")) < 1:
                driver.back()
                time.sleep(3)
            driver.refresh()
        increment += 1
    return False

def NavToUsers(driver,waitTime = 20):
    wait = WebDriverWait(driver,waitTime)
    
    for increment in range(6):
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@title = 'Home Menu']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text() = 'Settings']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Manage Users')]"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH,f"//td[text() = 'admin']")))
            return
        except:
            driver.refresh()
            time.sleep(1)
    return False

def CreateNewAdminUser(UserName, Email, Password, driver,waitTime = 20):
    wait = WebDriverWait(driver, waitTime)
        
    if len(driver.find_elements(By.XPATH,f"//*[text() = 'admin']")) > 0 and len(driver.find_elements(By.XPATH,f"//*[text() = 'Users']")) > 0:
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH,f"//*[text() = 'admin']"))).click()
            ActionBut = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Action')]")))
            ActionBut.click()
            #input("Creating account")
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Duplicate')]"))).click()
        except:
            print("failed to Create User because program failed to duplicate admin account")
            return False
        email = wait.until(EC.element_to_be_clickable((By.ID,"login")))
        while email.get_attribute("value") == "admin":
            time.sleep(1)
            
        try:
            name = wait.until(EC.element_to_be_clickable((By.ID,"name")))
            while name.get_attribute("value") != UserName:
                name.clear()
                name.send_keys(UserName)

            while email.get_attribute("value") != Email:
                email.clear()
                email.send_keys(Email)
            ActionBut.click()
        except:
            print("failed to Create User because program failed to enter Name or Email")
            return False
        
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Change Password')]"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//td[@name = 'new_passwd']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class = 'o_input' and @type = 'password']"))).send_keys(Password)
            wait.until(EC.element_to_be_clickable((By.NAME, "change_password_button"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text() = 'Users']"))).click()
        except:
            print("failed to Create User because program failed to change password or return to users")
            return False
        return True
    print("failed to Create User because program could not find Admin account or users")
    return False

def ChangePassword(Link, Email, Password, driver, waitTime = 20):
    wait = WebDriverWait(driver, waitTime)
    if len(driver.find_elements(By.XPATH,f"//*[text() = 'admin']")) > 0 and len(driver.find_elements(By.XPATH,f"//*[text() = 'Users']")) > 0:
        tries = 0
        while len(driver.find_elements(By.XPATH, "//*[text() = 'Action']")) < 1 and tries < 3:
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH,f"//td[text() = '{Email}']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH,"//*[text() = 'Action']"))).click()
            except ElementClickInterceptedException:    
                wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Filters')]"))).click()
            except:
                NavToUsers(driver,10)
                driver.refresh()
            tries += 1

        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Change Password')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//td[@name = 'new_passwd']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class = 'o_input' and @type = 'password']"))).send_keys(Password)
        wait.until(EC.element_to_be_clickable((By.NAME, "change_password_button"))).click()

        if Email == "admin":
            driver.get(Link)
            LoggedIn = Login(Creds["Emails"][1],Password,Link,driver,10)
            NavToUsers(driver,10)
            driver.refresh()
            if not LoggedIn:
                print("Failed to Login after changing admin password")
                return False
        else:
            driver.back()
            driver.refresh()
            time.sleep(1)
        return True
    print(f"Failed to Change Password for {Email}")
    return False

def Main(NewCreds:dict,Creds:dict,cred):
    Link = f"https://{cred.lower()}.ka-ching.asia/web/login"
    driver = webdriver.Chrome(options=ChromeOptions,service=ChromeService)
    driver.get(Link) 
    wait = WebDriverWait(driver, 20)
    #login to site
    LoggedIn = Login(Creds["Emails"][1],Creds[cred][1],Link,driver,10)
    if not LoggedIn:
        print(f"https://{cred.lower()}.ka-ching.asia/web/login Failed to login")
        driver.quit()
        return False
    NavToUsers(driver,10)

    try:
    #check if every account exists. if all accounts exist, change their passwords
        for Email in Creds["Emails"]:
            Password = NewCreds[cred][Creds["Emails"].index(Email)]
            wait.until(EC.element_to_be_clickable((By.XPATH,f"//*[text() = 'Filters']")))

            if len(driver.find_elements(By.XPATH,f"//td[text() = '{Email}']")) < 1:
                UserName = Creds["Instance Name"][Creds["Emails"].index(Email)]
                CreatedUser = CreateNewAdminUser(UserName,Email,Password,driver,10)
                if CreatedUser:
                    print(f"an account for {Email} was created at {cred.lower()}.ka-ching.asia")
                else:
                    driver.quit()
                    return False
            else:
                ChangePassword(Link,Email,Password,driver)
        print(f"Passwords changed and all accounts created at https://{cred.lower()}.ka-ching.asia")
        PassedCreds = []
        for Email in Creds["Emails"]:
            driver.get(Link)
            Password = NewCreds[cred][Creds["Emails"].index(Email)]
            LoggedIn = Login(Email,Password,Link,driver,10)
            if LoggedIn:
                driver.get(Link)
                PassedCreds.append(Email)
            else:
                print(f"{Email} failed to authenticate, failed to verify new creds")

        if len(PassedCreds) == len(Creds["Emails"]):
            print("All Credentials verified",end="")
        else:
            print("the following credentials have been verified ", PassedCreds)
        driver.quit()
        return True
    except Exception as e:
        print(f"Failed to change passwords or create accounts at https://{cred.lower()}.ka-ching.asia\n{e}")
        driver.quit()
        return False

#ChromeOptions.add_experimental_option("detach",True)
ChromeOptions.add_argument('--log-level=3')
ChromeOptions.add_argument('ignore-certificate-errors')
ChromeOptions.add_argument("--disable-extensions")
ChromeOptions.add_argument("--disable-gpu")
ChromeOptions.accept_insecure_certs = True
ChromeOptions.add_argument("--headless=old")
ChromeService = Service()
ChromeService.creation_flags = CREATE_NO_WINDOW

#requests for path to old password file, generates new passwords based on number of sites and users. new passwords is outputted to a file
#OldPasswords = input("please enter absolute File Path: ")
#NewCreds = GenNewPW(GetCredsFromExcel(OldPasswords))
#FilePath = DictToExcel(dic=NewCreds)
#print(f"New Passwords have been outputed to {FilePath}")

OldPasswords ="C:/Users/asher/OneDrive/Desktop/Intern Projects/Python/SeleniumPasswordUpdater/Files/NewPWs.csv"
start = time.time()
NewCreds = GetCredsFromExcel(OldPasswords)

#removes username and emails from new credentials dict
NewCreds.pop("Emails")
NewCreds.pop("Instance Name")

Creds = GetCredsFromExcel(OldPasswords)
total = len(NewCreds)
finished = 0
tries = 0

#loops through all instances/sites to change the passwords
while len(NewCreds) > 0 and tries < 3:
    retries = 0
    LoopSet = dict(NewCreds)
    for Site in LoopSet:
        Passed = Main(LoopSet,Creds,Site)
        #if the password change operation has no errors/exceptions, the instance/site is removed from the loop 
        if Passed:
            NewCreds.pop(Site)
            retries += 1
        print(f"({retries} out of {len(LoopSet)} passed)")
    tries += 1
    finished += retries
end = time.time()
print(f"this program has a {100 * finished / total}% success. ({finished} out of {total} passed)")
print(f"Failed sites {NewCreds}")
print(f"this program has run for {round((end - start) / 60)} minutes and {round((end - start) % 60)} seconds")