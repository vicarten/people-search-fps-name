#(A) Install packages
'''
!pip install selenium
!pip install undetected-chromedriver
!pip install webdriver-manager
!pip install pandas
!pip install chromedriver-autoinstaller
!pip install openpyxl
'''

#(B) Import libraries

from selenium import webdriver
import chromedriver_autoinstaller 
import pandas as pd
import openpyxl as xls
import time
import undetected_chromedriver as uc
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
#_________________________________________________________________________________________________________________________________________________________

#(C) Read excel file

fps = pd.read_excel("test.xlsx")
num_link = len(fps)
fps
#_________________________________________________________________________________________________________________________________________________________

#(D) Create a data frame

data_frame = pd.DataFrame(columns = ['row id', 'owner','provided', 'name', 'aka', 'age', 'address', 'past address', 'current address property details', 'primary phone', 'other phone numbers', 'emails', 'link', 'tbc'])
#_________________________________________________________________________________________________________________________________________________________

#(E) Define functions

#prepare address string for comparison
def edit(address):
    address = address.replace(" ", "")
    address = address.replace(".", "")
    return address.lower()

#scrape page info
def info():
    #name
    name = driver.find_element(By.XPATH,"//h1[@id='details-header']").text
    delimiter = "\n"
    index = name.index(delimiter)
    name = name[:index]

    #also known as
    try:
        outer_class = driver.find_element(By.XPATH,"//div[@id='aka-links']//div[@class='detail-box-email']")
        aka = outer_class.find_element(By.CLASS_NAME, "row").text
        aka = aka.replace("\n", " / ")
    except NoSuchElementException:
        aka = ""
        pass

    #age 
    try:
        age = driver.find_element(By.XPATH,"//h2[@id='age-header']").text
    except NoSuchElementException:
        age = ""
        pass

    #primary phone number
    try:
        phone = driver.find_element(By.XPATH,"//a[starts-with(@title,'Search people associated with the phone number')]").text
    except NoSuchElementException:
        phone = ""
        pass
    
    #other phone numbers
    try:
        phone_numbers = ""
        outer_class = driver.find_element(By.XPATH,"//div[@id='phone_number_section']")
        inner_class = outer_class.find_element(By.XPATH, "//div[@id='phone_number_section']//dl[1]")
        other_phones = inner_class.find_elements(By.XPATH, "//dl[@class='col-sm-12 col-md-6']")
        for other_phone in other_phones:
            other_phone = other_phone.text
            other_phone = other_phone.replace("\n", ", ")
            phone_numbers = phone_numbers + other_phone + " / "
    except NoSuchElementException:
        pass

    #emails
    try:
        outer_class = driver.find_element(By.XPATH,"//div[@id='email_section']//div[@class='detail-box-content']")
        email = outer_class.find_element(By.CLASS_NAME, "row").text
        email = email.replace("\n", " / ")
    except NoSuchElementException:
        email = ""
        pass
    
    #current address' property details
    try:
        all_details=""
        outer_class = driver.find_element(By.XPATH,"//div[@id='current_property_data']")
        details = outer_class.find_elements(By.CSS_SELECTOR, 'dl')
        for detail in details:
            detail = detail.text
            detail = detail.replace("\n", ": ")
            all_details = all_details + detail + " / "
    except NoSuchElementException:
        pass

    return name, aka, age, phone, phone_numbers, email, all_details
#_________________________________________________________________________________________________________________________________________________________

#(F) Set chrome options

options = uc.ChromeOptions()

prefs = {"credentials_enable_service": False,"profile.password_manager_enabled": False}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-notifications")
options.add_argument("--disable-popup-blocking") 

driver = uc.Chrome(options=options)
driver.maximize_window()
#_________________________________________________________________________________________________________________________________________________________

#(G) Set the search range

start = 0 #start line
num_link = 100 #end line
#_________________________________________________________________________________________________________________________________________________________

#(H) Iterate over each row in the range from 'test.xlsx'

for row in range(start, num_link): 
    print(row)
    id=row+2
    owner = fps.iloc[row,0]
    provided_address = fps.iloc[row,1]
    new_address = edit(str(provided_address))
    
   
    #default values
    match = False
    current_match = False
    current_maybe = False
    maybe = False
    name = ""
    phone = ""
    email = ""
    bot = ""
    save_current_address = ""
    save_past_address = ""
    
    #start the search with premade link from excel 
    link = fps.iloc[row,3]
    driver.get(link)
    time.sleep(1)
    
    #check if the page asks to complete CAPTCHA
    try:
        bot = driver.find_element(By.XPATH, "//h1[normalize-space()='Are you human?']").text
    except NoSuchElementException:
        pass
    if (bot == "Are you human?"):
        time.sleep(30)
 
    #check the current and past addresses
    if len(new_address)>10:
        new_address = new_address[:10]
    try:
        same_names=driver.find_elements(By.XPATH,"//span[@class='larger']")
        i=0
        # Iterate over the same_names elements
        for i in range(1, len(same_names) + 1):
            time.sleep(1)
           
            element = wait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"(//span[@class='larger'])[{str(i)}]")))
            driver.execute_script("arguments[0].scrollIntoView();", element)
            driver.execute_script("window.scrollBy(0, -100);")
            ActionChains(driver).move_to_element(element).click().perform()  
            time.sleep(1)
            
            #check current address
            try:
                current_address = driver.find_element(By.XPATH,"//a[starts-with(@title,'Search people living at')]").text
                save_current_address = current_address.replace("\n", ", ")
                current_address = edit(str(current_address))[:len(new_address)]
                if current_address == new_address:
                    current_match = True
                    match = True
                    name, aka, age, phone, phone_numbers, email, all_details = info()
                    new_row = [id, owner, provided_address, name, aka, age, save_current_address, "", all_details, phone, phone_numbers, email, link, ""]
                    data_frame.loc[len(data_frame)] = new_row
                if match == False and current_maybe == False and current_address[:4] == new_address[:4] and new_address[:4] != "pobo":
                    maybe = True
                    current_maybe = True
                    name, aka, age, phone, phone_numbers, email, all_details = info()
                    new_row = [id, owner, provided_address, name, aka, age, save_current_address, "", all_details, phone, phone_numbers, email, link, "!!!"]
                    data_frame.loc[len(data_frame)] = new_row
                    
            except NoSuchElementException:
                pass
            
            #check past addresses
            if match == False:
                try:
                    past_address = driver.find_elements(By.XPATH,"//a[starts-with(@title,'Search people who live at')]")
                    for p in range(1, len(past_address) + 1):
                        past = wait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, f"(//a[starts-with(@title,'Search people who live at')])[{str(p)}]")))
                        past=past.text
                        save_past_address = past.replace("\n", ", ")
                        past = edit(str(past))[:len(new_address)]
                        if past == new_address:
                            match = True
                            name, aka, age, phone, phone_numbers, email, all_details = info()
                            new_row = [id, owner, provided_address, name, aka, age, save_current_address, save_past_address, "", phone, phone_numbers, email, link, ""]
                            data_frame.loc[len(data_frame)] = new_row
                        if match == False and maybe == False and past[:4] == new_address[:4] and new_address[:4] != "pobo":
                            maybe = True
                            name, aka, age, phone, phone_numbers, email, all_details = info()
                            new_row = [id, owner, provided_address, name, aka, age, save_current_address, save_past_address, "", phone, phone_numbers, email, link, "!!!"]
                            data_frame.loc[len(data_frame)] = new_row
                        if p==12 or match==True:
                            break
                            
                except NoSuchElementException and TimeoutException:
                    pass
            time.sleep(1)    
            driver.back()    
            i+=1
            if i==5 or current_match == True:
                break
            
                
        if match == False and maybe == False:
            new_row = [id, owner, provided_address, "", "", "", "", "", "", "", "", "", "", ""]
            data_frame.loc[len(data_frame)] = new_row

    except NoSuchElementException:
        pass
    print("__________________" + str(row) + " complete")
#_________________________________________________________________________________________________________________________________________________________
    
#(I) Display the first and last five search results

display(data_frame)
#_________________________________________________________________________________________________________________________________________________________

#(J) Export to 'output.xlsx' file

#define the row highlighting function
def highlight_row(row):
    color = 'background-color: #CCABD8' if row[-1] == '!!!' else ''
    highlight_cols = ['background-color: #CFE2F3' if col in ['owner', 'provided'] else color for col in data_frame.columns]
    return highlight_cols

#apply the highlighting function to the data frame
styled_df = data_frame.style.apply(highlight_row, axis=1)

#display the styled data frame
styled_df

#save the styled data frame to excel
styled_df.to_excel('output.xlsx', index=False)
#_________________________________________________________________________________________________________________________________________________________