#This script was written by Richard Passett and downloads the daily csv files from FLPA and Grants Portal that we need for reporting and analysis. 
#It currently grabs 14 files.

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
from datetime import date

#Password variables for FLPA and Grants Portal
FLPA_username=keyring.get_password("FLPA_GP", "username")
FLPA_password=keyring.get_password("FLPA", "Passett")
GP_username=keyring.get_password("GP", "Reports username")
GP_password=keyring.get_password("GP", "Reports password")

#Directories used for script
holding_dir=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Holding Folder'
accounts_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Accounts Export'
appeals_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Appeals Export'
large_project_closeout_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Large Project Closeout Export'
project_amendments_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Project Amendments Export'
project_version_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Project Version Export'
projects_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Projects Export'
receivables_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Receivables Export'
extensions_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Time Extensions Export'
gp_full_subrecipient_projects_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal Full Subrecipient Projects Export'
gp_full_fdem_projects_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal Full FDEM Projects Export'
gp_active_subrecipient_projects_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal Active Subrecipient Projects Export'
gp_active_fdem_projects_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal Active FDEM Projects Export'
gp_EEI_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal EEI Export'
gp_RFI_destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Grants Portal RFI Export'

#FLPA Report Listings used for script
Accounts_Listing="https://floridapa.org/app/#account/accountlist?o=grantname+asc%2Capplicantname+asc"
Appeals_Listing="https://floridapa.org/app/#project/projectappeallist?"
large_project_closeout_listing="https://floridapa.org/app/#project/projectcloseoutlist?filters=%7B%22Program%22%3A%221%22%2C%22Step%22%3A%2226%2C27%2C28%2C570%2C29%2C482%2C485%2C157%2C571%2C183%2C446%2C572%2C159%22%7D&pp=25&o=laststepchangedays+asc"
Project_Amendments_Listing="https://floridapa.org/app/#project/projectscopechangelist?o=laststepchangedays+asc"
Project_Version_Listing="https://floridapa.org/app/#project/projectversionlist?filters=%7B%22Step%22%3A%22123%2C509%22%7D&o=laststepchangedays+asc&p=1&pp=50&s="
Project_Listing="https://floridapa.org/app/#project/projectlist?o=programshortname+asc"
Receivables_Listing="https://floridapa.org/app/#payment/receivablelist?"
Extensions_Listing="https://floridapa.org/app/#project/projectextensionlist?o=laststepchangedays+asc"

#Grants Portal Listings used for script
GP_Full_Subrecipient_Projects_Listing="https://grantee.fema.gov/#projects/subrecipient?filters=1668473"
GP_Full_FDEM_Projects_Listing="https://grantee.fema.gov/#projects?filters=1669612"
GP_Active_Subrecipient_Projects_Listing="https://grantee.fema.gov/#projects/subrecipient?filters=1668412"
GP_Active_FDEM_Projects_Listing="https://grantee.fema.gov/#projects?filters=1669611"
GP_EEI_Listing="https://grantee.fema.gov/#projectEEI/subrecipient?filters=1669615"
GP_RFI_Listing="https://grantee.fema.gov/#rfis/subrecipient?filters=1669613"

#Use webdriver for Chrome, set where you want the CSVs to download to, add other options/preferences as desired, point to where you have the driver downloaded, and set the driver to a variable.
#If you want to see what is happening in the browser, comment out the headless and disable-software-rasterizer options
options=webdriver.ChromeOptions()
prefs={"download.default_directory" : r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Holding Folder'}
options.add_experimental_option("prefs",prefs) 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--headless")
options.add_argument("--disable-software-rasterizer")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 120)

#Function that downloads CSV files from FLPA.
#The process is the same with the same locations for all reports, which is why we can build a reusable function for this.
def download_FLPA_report():
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

def download_GP_report():
    wait.until(EC.presence_of_element_located((By.CLASS_NAME,'caret')))
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'caret')))
    dropdown_button=driver.find_element(By.CLASS_NAME,'caret')
    driver.execute_script("arguments[0].click();", dropdown_button)
    time.sleep(3)
    export_button=driver.find_element(By.XPATH,'//*[@id="accordion"]/div/div[1]/div[2]/div[2]/div/ul/li[5]/a')
    driver.execute_script("arguments[0].click();", export_button)

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

#Function to rename export file
def Rename_File(folder, file_name):
    for file in os.listdir(folder):
        old_file_name=folder+"/"+file
        if file.endswith(".csv"):
            new_file_name=folder+"/"+file_name+date.today().strftime("%m%d%Y")+".csv"
            os.rename(old_file_name, new_file_name)
        elif file.endswith(".xlsx"):
            new_file_name=folder+"/"+file_name+date.today().strftime("%m%d%Y")+".xlsx"
            os.rename(old_file_name, new_file_name)
        else:
            return

#Function to download FLPA CSVs. Accepts two arguments; listing (driver.get location) and destination (destination directory) 
def FLPA_export(listing, destination, name):
    driver.get(listing)
    time.sleep(5)
    download_FLPA_report()
    move(destination)
    Rename_File(destination, name)

#Function to Export data from Grants Portal in CSV format. Accepts two arguments; listing (driver.get location) and destination (destination directory) 
def GP_export(listing, destination, name):
    driver.get(listing)
    time.sleep(5)
    download_GP_report()
    move(destination)
    Rename_File(destination, name)

#Provide a message to the person running the script
print("Greetings, we are updating your 14 csv files to reflect the latest data from FLPA and Grants Portal.\nThis process will take about 25 minutes with the majority of that time used getting the FLPA Projects Export.\nWe will provide status updates along the way.")

#Open FLPA
driver.get("https://floridapa.org/")
time.sleep(8)

#Login to FLPA
username_field=driver.find_element(By.NAME,"Username")
password_field=driver.find_element(By.NAME,"Password")
signIn_button=driver.find_element(By.NAME,"Submit")
username_field.clear()
password_field.clear()
username_field.send_keys(FLPA_username)
password_field.send_keys(FLPA_password)
signIn_button.click()
time.sleep(5)

#Export FLPA Data
FLPA_export(Accounts_Listing, accounts_destination, "FLPA_Accounts_Export_")
print("1st task complete. Your Accounts data is located here: "+accounts_destination)

FLPA_export(Appeals_Listing, appeals_destination, "FLPA_Appeals_Export_")
print("2nd task complete. Your Appeals data is located here: "+appeals_destination)

FLPA_export(large_project_closeout_listing, large_project_closeout_destination, "FLPA_Large_Project_Closeout_Export_")
print("3rd task complete. Your Large Project Closeout data is located here: "+large_project_closeout_destination)

FLPA_export(Project_Amendments_Listing, project_amendments_destination, "FLPA_Project_Amendments_Export_")
print("4th task complete. Your Project Amendment data is located here: "+project_amendments_destination)

FLPA_export(Project_Version_Listing, project_version_destination, "FLPA_Project_Version_Export_")
print("5th task complete. Your Project Version data is located here: "+project_version_destination)

FLPA_export(Project_Listing, projects_destination, "FLPA_Project_Export_")
print("6th task complete. Your Project data is located here: "+projects_destination)

FLPA_export(Receivables_Listing, receivables_destination, "FLPA_Receivables_Export_")
print("7th task complete. Your Receivables data is located here: "+receivables_destination)

FLPA_export(Extensions_Listing, extensions_destination, "FLPA_Extensions_Export_")
print("8th task complete. Your Time Extension data is located here: "+extensions_destination)

#The rest of the code below is used to grab our Grants Portal data. 
 
#Open Grants Portal
driver.get("https://grantee.fema.gov/")
time.sleep(15)

#Login to Grants Portal
username_field=driver.find_element(By.ID,"username")
password_field=driver.find_element(By.ID,"password")
signIn_button=driver.find_element(By.ID,"credentialsLoginButton")
username_field.clear()
password_field.clear()
username_field.send_keys(GP_username)
password_field.send_keys(GP_password)
signIn_button.click()
time.sleep(15)
accept_button=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
accept_button.click()
time.sleep(10)
accept_button2=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
time.sleep(10)
accept_button2.click()
time.sleep(15)

#Export Grants Portal Data
GP_export(GP_EEI_Listing, gp_EEI_destination, "EEI_Export_")
print("9th task complete. Your GP EEI data is located here: "+gp_EEI_destination)

GP_export(GP_RFI_Listing, gp_RFI_destination, "RFI_Export_")
print("10th task complete. Your GP RFI data is located here: "+gp_RFI_destination)

GP_export(GP_Active_Subrecipient_Projects_Listing, gp_active_subrecipient_projects_destination, "GP_Active_Subrecipient_Projects_Export_")
print("11th task complete. Your GP Active Subrecipient Project data is located here: "+gp_active_subrecipient_projects_destination)

GP_export(GP_Active_FDEM_Projects_Listing, gp_active_fdem_projects_destination, "GP_Active_FDEM_Projects_Export_")
print("12th task complete. Your GP Active FDEM Project data is located here: "+gp_active_fdem_projects_destination)

GP_export(GP_Full_FDEM_Projects_Listing, gp_full_fdem_projects_destination, "GP_Full_FDEM_Projects_Export_")
print("13th task complete. Your GP Full FDEM Project data is located here: "+gp_full_fdem_projects_destination)

GP_export(GP_Full_Subrecipient_Projects_Listing, gp_full_subrecipient_projects_destination, "GP_Full_Subrecipient_Projects_Export_")
print("14th and final task complete! Your GP Full Subrecipient Project data is located here: "+gp_full_subrecipient_projects_destination)

driver.close()
