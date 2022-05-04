#This script was written by Richard Passett and downloads the daily csv files from FLPA and Grants Portal that we need for reporting and analysis. 
#It currently grabs 8 files.

#Import dependencies
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from zipfile import ZipFile
import os
import shutil
import keyring

#Password variables for FLPA and Grants Portal
my_username=keyring.get_password("FLPA_GP", "username")
FLPA_password=keyring.get_password("FLPA", "Passett")
GP_password=keyring.get_password("GP", "Passett")

#Directories used for script
holding_dir=r'C:\Users\richardp\Desktop\FLPA_CSVs'
accounts_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Accounts Export'
appeals_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Appeals Export'
project_amendments_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Project Amendments Export'
projects_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Projects Export'
receivables_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Receivables Export'
extensions_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Time Extensions Export'
grants_portal_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\Grants Portal Projects Export'
project_version_destination=r'J:\Admin & Plans Unit\PA Reporting Section\Richard Tests\FLPA Project Version Export'

#Report Listings used for script
Accounts_Listing="https://floridapa.org/app/#account/accountlist?o=grantname+asc%2Capplicantname+asc"
Appeals_Listing="https://floridapa.org/app/#project/projectappeallist?"
Project_Amendments_Listing="https://floridapa.org/app/#project/projectscopechangelist?o=laststepchangedays+asc"
Project_Listing="https://floridapa.org/app/#project/projectlist?o=programshortname+asc"
Receivables_Listing="https://floridapa.org/app/#payment/receivablelist?"
Extensions_Listing="https://floridapa.org/app/#project/projectextensionlist?o=laststepchangedays+asc"
GP_Projects_Listing="https://grantee.fema.gov/#projects/subrecipient?filters=1581675"
Project_Version_Listing="https://floridapa.org/app/#project/projectversionlist?filters=%7B%22Step%22%3A%22123%2C509%22%7D&o=laststepchangedays+asc&p=1&pp=50&s="

#Use webdriver for Chrome, set where you want the CSVs to download to, add other options/preferences as desired, point to where you have the driver downloaded, and set the driver to a variable.
#If you want to see what is happening in the browser, comment out the headless and disable-software-rasterizer options
options=webdriver.ChromeOptions()
prefs={"download.default_directory" : r"C:\Users\richardp\Desktop\FLPA_CSVs"}
options.add_experimental_option("prefs",prefs) 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--headless")
options.add_argument("--disable-software-rasterizer")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 120)

#Function that downloads CSV files from FLPA.
#The process is the same with the same locations for all reports, which is why we can build a reusable function for this.
def download_report():
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.toExcel.inner")))
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.toExcel.inner")))
    Excel_button=driver.find_element(By.CSS_SELECTOR,"div.toExcel.inner")
    driver.execute_script("arguments[0].click();", Excel_button)
    wait.until(EC.element_to_be_clickable((By.ID,'excelexportcolumns2')))
    Custom_button=driver.find_element(By.ID,'excelexportcolumns2')
    driver.execute_script("arguments[0].click();", Custom_button)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'selectAll')))
    selectAll_button=driver.find_element(By.CLASS_NAME,'selectAll')
    driver.execute_script("arguments[0].click();", selectAll_button)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input.close.main")))
    export_button=driver.find_element(By.CSS_SELECTOR,"input.close.main")
    driver.execute_script("arguments[0].click();", export_button)
    time.sleep(2)
    try:
        driver.find_elements(By.CSS_SELECTOR,"input.close.main")[-1].click()
    except IndexError:
        pass

#function to move csv to desired destination. Waits for file to exist, empties destination folder before moving new file, and accounts for whether or not csv is in a zip file.
def move(destination):
    while len(os.listdir(holding_dir))==0: 
        time.sleep(10)
    for file in os.scandir(destination):
        os.remove(file.path)
    for item in os.listdir(holding_dir):
        file_name=holding_dir+"/"+item
        if item.endswith(".zip"):
            zip_ref = ZipFile(file_name) # create zipfile object
            zip_ref.extractall(destination) # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) #Delete original file
        elif item.endswith("crdownload"):
            time.sleep(10)
            move(destination)
        else:
            shutil.copy2(file_name, destination) #Copy csv to JDrive
            os.remove(file_name) #Delete original file
    time.sleep(5)

#Function to download FLPA CSVs. Accepts two arguments; listing (driver.get location) and destination (destination directory) 
def export(listing, destination):
    driver.get(listing)
    time.sleep(3)
    download_report()
    move(destination)

#Provide a message to the person running the script
print("Greetings, we are updating your 8 csv files to reflect the latest data from FLPA and Grants Portal.\nThis process will take about 20 minutes with the majority of that time used on task 4 (FLPA Projects Export).\nWe will provide status updates along the way.")

#Open FLPA
driver.get("https://floridapa.org/")
time.sleep(8)

#Login to FLPA
username_field=driver.find_element(By.NAME,"Username")
password_field=driver.find_element(By.NAME,"Password")
signIn_button=driver.find_element(By.NAME,"Submit")
username_field.clear()
password_field.clear()
username_field.send_keys(my_username)
password_field.send_keys(FLPA_password)
signIn_button.click()
time.sleep(5)

#Export FLPA Reports
export(Accounts_Listing, accounts_destination)
print("1st task complete. Your Accounts data is located here: "+accounts_destination)

export(Appeals_Listing, appeals_destination)
print("2nd task complete. Your Appeals data is located here: "+appeals_destination)

export(Project_Amendments_Listing, project_amendments_destination)
print("3rd task complete. Your Project Amendment data is located here: "+project_amendments_destination)

export(Project_Listing, projects_destination)
print("4th task complete. Your Project data is located here: "+projects_destination)

export(Receivables_Listing, receivables_destination)
print("5th task complete. Your Receivables data is located here: "+receivables_destination)

export(Extensions_Listing, extensions_destination)
print("6th task complete. Your Time Extension data is located here: "+extensions_destination)

export(Project_Version_Listing, project_version_destination)
print("7th task complete. Your Project Version data is located here: "+project_version_destination)

#The rest of the code below is used to grab our Grants Portal data. 
#I will write a function for exporting csvs from Grants Portal (similar to the download_report function for FLPA data) whenever I need to download more than one report.
 
#Open Grants Portal
driver.get("https://grantee.fema.gov/")
time.sleep(15)

#Login to Grants Portal
username_field=driver.find_element(By.ID,"username")
password_field=driver.find_element(By.ID,"password")
signIn_button=driver.find_element(By.ID,"credentialsLoginButton")
username_field.clear()
password_field.clear()
username_field.send_keys(my_username)
password_field.send_keys(GP_password)
signIn_button.click()
time.sleep(10)
accept_button=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
accept_button.click()
time.sleep(4)
accept_button2=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
time.sleep(4)
accept_button2.click()
time.sleep(8)

#Export Subrecipient Projects
driver.get("https://grantee.fema.gov/#projects/subrecipient?filters=1581675")
wait.until(EC.presence_of_element_located((By.CLASS_NAME,'caret')))
wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'caret')))
dropdown_button=driver.find_element(By.CLASS_NAME,'caret')
driver.execute_script("arguments[0].click();", dropdown_button)
time.sleep(3)
export_button=driver.find_element(By.XPATH,'/html/body/div[2]/div[1]/div/div/div[3]/div[2]/div/div[2]/div/div/div/div/div[1]/div[2]/div[2]/div/ul/li[5]/a')
driver.execute_script("arguments[0].click();", export_button)
move(grants_portal_destination)

driver.close()
print("8th and final task complete! Your Grants Portal data has been saved here: "+grants_portal_destination)