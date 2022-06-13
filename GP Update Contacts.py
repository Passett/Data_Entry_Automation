#This script was written by Richard Passett
#Logs into Grants Portal and updates all SPAO and Alternate SPAO for all subrecipients at the Subrecipient Applicant Event Profile Level

#Import dependencies
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import keyring
import pandas as pd

#destination for CSVs of link lists
folder=r'C:\Users\richardp\Desktop\Contact Automation Tests'

#Define SPAO and Alternate placeholders for each subrecipient. Use people that you know will be available options in the dropdown list of contacts
#Define the actual SPAO and Alternate SPAO you want to update subrecipients with

SPAO_Placeholder="passett"
Alternate_Placeholder="blocker"
Current_SPAO="adkison"
Current_Alternate_SPAO="morales"

#List of links to applicant event profiles that need to be updated
df_list=pd.read_excel(r'C:\Users\richardp\Desktop\Contact Automation Tests\Final_List.xlsx',sheet_name=0)
subrecipient_links=df_list['Applicant_Url'].tolist()

#Lists that will get filled and sent to CSVs
processed=[]
need_contacts=[]
skipped=[]

#Function to update contacts at Applicant Event Profile Level
def single_page_update_Contacts(url):
    outer_error_counter=0
    driver.get(url)
    time.sleep(5)
    try:
        wait.until(EC.element_to_be_clickable((By.ID,'manage-contacts-button')))
        Contacts_button=driver.find_element(By.ID,'manage-contacts-button')
        driver.execute_script("arguments[0].click();", Contacts_button) #Clicks the manage POCS button
        time.sleep(5)
        update_recipients(url)

    except:
        outer_error_counter+=1
        if outer_error_counter>2:
            print("error occurred while processing "+url)
            skipped.append(url)
        else:
            driver.refresh()
            time.sleep(18)
            wait.until(EC.element_to_be_clickable((By.ID,'manage-contacts-button')))
            Contacts_button=driver.find_element(By.ID,'manage-contacts-button')
            driver.execute_script("arguments[0].click();", Contacts_button) #Clicks the manage POCS button
            time.sleep(8)
            update_recipients(url)

    '''update recipient function below checks to see if the expected four options are available for selection in the "Manage Applicant Event Profile Contacts" popup
    If they are, it next checks to make sure they have a primary contact. 
    If they do, it moves on. If they do not, it appends the url to the need_contacts list, hits the cancel button and moves to next url
    If they aren't missing a primary contact, it clicks the down arrow next to recipient POC and puts a place holder in for the contact
    Then it does the same thing for the alternate recipient
    Placeholders are used because if a contact is already used in one field, they will not be available in other
    Often times, if your recipient leaves your alternate will replace them
    The placeholders are used to allow the promoted alternate to be included in the list options for the recipient
    After putting in place holders, it then replaces them with the actual contacts
    If for some reason, the expected four options are not available in the popup, 
    it will click the cancel button, append the url to the skipped list, and print a note stating that the applicant event profile was not updated
    if it times out trying to click the manage POCs button, it refreshes the page and tries again.
    It will make this refresh attempt twice before skipping the record and appending it to the skipped list '''


def update_recipients(url):
    error_counter=0
    Contact_options=driver.find_elements(By.CLASS_NAME,"selection")
    try:
        if len(Contact_options)==4:
            
            if "This field is required." in driver.find_element(By.CLASS_NAME,"has-error").text:
                print("Primary contact is missing for "+url)
                need_contacts.append(url)
                wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[2]')))
                Cancel_button=driver.find_element(By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[2]')
                Cancel_button.click()

            else:
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME,"select2-selection__arrow")))
                SPAO_Down_Arrow=driver.find_elements(By.CLASS_NAME,"select2-selection__arrow")[2]
                SPAO_Down_Arrow.click()
                SPAO_Element=driver.find_element(By.CLASS_NAME,"select2-search__field")
                SPAO_Element.send_keys(SPAO_Placeholder)
                time.sleep(1)
                SPAO_Element.send_keys(Keys.ENTER)
                time.sleep(1)

                wait.until(EC.element_to_be_clickable((By.CLASS_NAME,"select2-selection__arrow")))
                Alternate_SPAO_Down_Arrow=driver.find_elements(By.CLASS_NAME,"select2-selection__arrow")[3]
                Alternate_SPAO_Down_Arrow.click()

                Alternate_SPAO_Element=driver.find_element(By.CLASS_NAME,"select2-search__field")
                Alternate_SPAO_Element.send_keys(Alternate_Placeholder)
                time.sleep(1)
                Alternate_SPAO_Element.send_keys(Keys.ENTER)
                time.sleep(1)

                wait.until(EC.element_to_be_clickable((By.CLASS_NAME,"select2-selection__arrow")))
                SPAO_Down_Arrow=driver.find_elements(By.CLASS_NAME,"select2-selection__arrow")[2]
                SPAO_Down_Arrow.click()

                SPAO_Element=driver.find_element(By.CLASS_NAME,"select2-search__field")
                SPAO_Element.send_keys(Current_SPAO)
                time.sleep(1)
                SPAO_Element.send_keys(Keys.ENTER)
                time.sleep(1)

                wait.until(EC.element_to_be_clickable((By.CLASS_NAME,"select2-selection__arrow")))
                Alternate_SPAO_Down_Arrow=driver.find_elements(By.CLASS_NAME,"select2-selection__arrow")[3]
                Alternate_SPAO_Down_Arrow.click()

                Alternate_SPAO_Element=driver.find_element(By.CLASS_NAME,"select2-search__field")
                Alternate_SPAO_Element.send_keys(Current_Alternate_SPAO)
                time.sleep(1)
                Alternate_SPAO_Element.send_keys(Keys.ENTER)
                time.sleep(1)

                wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[1]')))
                Save_button=driver.find_element(By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[1]')
                Save_button.click()
                print(url+" has been successfully updated!")
                processed.append(url)
                time.sleep(8)

        else:
            wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[2]')))
            Cancel_button=driver.find_element(By.XPATH,'//*[@id="MainWrapper"]/div[3]/div/div/div[3]/button[2]')
            Cancel_button.click()
            print(url+" not updated")
            skipped.append(url)
            time.sleep(5)

    except:
        error_counter+=1
        if error_counter>2:
            print("error occurred while processing "+url)
            skipped.append(url)
        else:
            driver.refresh()
            time.sleep(18)
            wait.until(EC.element_to_be_clickable((By.ID,'manage-contacts-button')))
            Contacts_button=driver.find_element(By.ID,'manage-contacts-button')
            driver.execute_script("arguments[0].click();", Contacts_button) #Clicks the manage POCS button
            time.sleep(8)
            update_recipients(url)

#Password variables for Grants Portal
my_username=keyring.get_password("FLPA_GP", "username")
GP_password=keyring.get_password("GP", "Passett")

#Use webdriver for Chrome and add other options/preferences as desired, point to where you have the driver downloaded, and set the driver to a variable.
#If you want to see what is happening in the browser, comment out the headless and disable-software-rasterizer options
options=webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument("--headless")
#options.add_argument("--disable-software-rasterizer")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 35)

#Open Grants Portal
driver.get("https://grantee.fema.gov/")
driver.maximize_window()
time.sleep(8)

#Login to Grants Portal
username_field=driver.find_element(By.ID,"username")
password_field=driver.find_element(By.ID,"password")
signIn_button=driver.find_element(By.ID,"credentialsLoginButton")
username_field.clear()
password_field.clear()
username_field.send_keys(my_username)
password_field.send_keys(GP_password)
signIn_button.click()
time.sleep(5)
accept_button=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
accept_button.click()
time.sleep(5)
accept_button2=driver.find_element(By.CSS_SELECTOR,"button.btn.btn-sm.btn-primary")
time.sleep(5)
accept_button2.click()
time.sleep(5)

#Loop over your list of subrecipient links and update the contacts for each one
for i in subrecipient_links:
    single_page_update_Contacts(i)

df1 = pd.DataFrame(processed, columns=["Links"])
df2 = pd.DataFrame(need_contacts, columns=["Links"])
df3 = pd.DataFrame(skipped, columns=["Links"])
df1.to_csv(folder+'\\processed.csv', index=False)
df2.to_csv(folder+'\\need_contacts.csv', index=False)
df3.to_csv(folder+'\\not_processed.csv', index=False)

Updates=len(subrecipient_links)
print("Task complete, "+Updates+" updates were processed!")