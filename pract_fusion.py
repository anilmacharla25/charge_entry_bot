from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import shutil
from datetime import datetime
import openpyxl
import tkinter as tk
from tkinter import messagebox
import os, pyautogui, pymsgbox
from selenium.webdriver.edge.options import Options

today_date = datetime.now().strftime('%m-%d-%Y')
# edge_driver_path = r"C:\Automation\Python Automation\500 Charge Entry BOT\msedgedriver.exe"
uid='username'
pwd='XXXXXXXX'

# driver = webdriver.Edge()
# driver.implicitly_wait(10)
def start_driver():
    global driver
    
    # os.system("start cmd")
    # time.sleep(2)
    # pyautogui.typewrite(r'cd C:\Program Files (x86)\Microsoft\Edge\Application')
    # pyautogui.hotkey('enter')
    # pyautogui.write(' msedge.exe --remote-debugging-port=9222')
    # pyautogui.hotkey('enter')
    os.system("start cmd") 
    time.sleep(2)

    pyautogui.typewrite(r'cd C:\Program Files (x86)\Microsoft\Edge\Application')
    pyautogui.hotkey('enter')

    pyautogui.write(r'"msedge.exe" ')
    pyautogui.write(' -')
    pyautogui.write('remote-debugging-port=9222')
    pyautogui.hotkey('enter')
    time.sleep(5)
    pyautogui.typewrite(r"https://static.practicefusion.com/apps/ehr/index.html#/login") ##make sure to keep only one edge browser instance to run in debug mode
    pyautogui.hotkey('enter')
    time.sleep(10)
    pyautogui.typewrite(r'XXXXXXXX')
    pyautogui.hotkey('enter')
    time.sleep(10)

    #connection
    # ------------ Connection -----------------------
    edge_options = Options()
    edge_options.use_chromium = True
    edge_options.add_experimental_option("debuggerAddress", "localhost:9222")

    driver = webdriver.Edge(options=edge_options)
    # pymsgbox.alert("Connected!")

def ask_for_security_code():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    # Show a confirmation dialog
    response = messagebox.askyesno("Confirmation", f"Did you type the security code?")
    if response:
        print("User confirmed. Proceeding...")
        return True
    else:
        return False
 
def ask_for_security_code1():
    driver.execute_script("alert('Did you type the security code?');")
    try:
        WebDriverWait(driver, 15).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()  # Accept the alert
        print("User confirmed. Proceeding...")
        return True, None  # User confirmed, no error
    except Exception as e:
        print(f"User did not confirm. Waiting for 15 seconds... Error: {e}")
        return False, e  # User did not confirm, return the exception

def remind_user():
    while True:
        # Display a JavaScript confirm alert
        script = '''
            var result = confirm('Click on OK if security code is typed, Otherwise Ignore');
            return result
        '''
        confirmation = driver.execute_script(script)
        time.sleep(7)

        if confirmation:
            print("User confirmed. Proceeding...")
            break
        else:
            print("User did not confirm. Waiting for 15 seconds...")
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            driver.switch_to.alert.dismiss()
            time.sleep(10)








def login_practice_fusion(uid, pwd):
    # website_url = r"https://static.practicefusion.com/apps/ehr/index.html#/login"
    # driver.get(website_url)
    start_driver()
    driver.maximize_window()
    time.sleep(5)
    user_name  = driver.find_element(By.ID, 'inputUsername')
    user_name.send_keys(uid)
    password = driver.find_element(By.ID, "inputPswd")
    password.send_keys(pwd)
    submit_btn = driver.find_element(By.ID,'loginButton')
    submit_btn.click()
    element = WebDriverWait(driver, 10) .until(EC.presence_of_element_located((By.ID, 'sendCallButton')))
    element.click()
    time.sleep(180)
    submit_code = WebDriverWait(driver, 10) .until(EC.presence_of_element_located((By.ID, 'sendCodeButton')))
    submit_code.click()
    time.sleep(2)
            



def show_bills_windows():
    try:
        menu_bar=driver.find_element(By.XPATH, "//button[contains(@class, 'menu-toggle') and contains(@class, 'nav-button') and @type='button']")
        menu_bar.click()
    except:
        pass
    reports =driver.find_element(By.CLASS_NAME,'reports')
    reports.click()
    billing_report = driver.find_element(By.XPATH,"//h2[text()='Practice management']/following-sibling::div//a[contains(text(), 'Billing report')]")                                            
    driver.execute_script("arguments[0].scrollIntoView();", billing_report)
    billing_report.click()


def scroll_to_element(element):
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()


def get_bill_links():
    bill_id_links=[]
    count=1
    while True:
        try:
            row = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,f'//tr[@aria-rowindex="{count}"]'))
            )
        
            # Scroll to the element using ActionChains
            scroll_to_element(row)
            time.sleep(2)
            print(count)
            print([cell.text for cell in row.find_elements(By.TAG_NAME,'td')])
            cells=[cell for cell in row.find_elements(By.TAG_NAME, 'td')]
            if len(cells)==0:
                continue
            elif cells[-1].text.strip()=='Draft':
                print(cells[0].text)
                bill_link= cells[0].find_element(By.TAG_NAME,'a').get_attribute('href')
                
                bill_id_links.append(bill_link)   
        except:
            break
        count+=1
    return bill_id_links


def get_service_data(service_number):
    #service data- CPT and ICDs and start and end dates of service
    try:
        service_div= driver.find_element(By.XPATH, f"//div[@data-element='superbill-service-{service_number}']")
    except NoSuchElementException:
        return None
    service_cpt= service_div.find_element(By.TAG_NAME,"h4").text
    icds_div=service_div.find_element(By.TAG_NAME,"fieldset")
    service_icds= [icd.text for icd in icds_div.find_elements(By.TAG_NAME,'h4')]
    icds_text= ','.join(service_icds)
    print(service_cpt)

    inputs_list=service_div.find_elements(By.TAG_NAME,"input")
    dates_list= []
    for input in inputs_list:
        if len(input.get_attribute('value'))<7:
            continue
        else:
            dates_list.append(input.get_attribute('value'))
    if len(dates_list)==0:
        dates_data= driver.find_element(By.XPATH, "//p[@class='p' and @data-element='start-to-end-date-read-only']")
        dates_list= dates_data.text.strip().split('-')
    return service_cpt,icds_text,dates_list[0],dates_list[1]
  

def get_pt_details():
    #fetching pt details
    try:
        bill_id=driver.find_element(By.XPATH,"//span[@class='text-color-default' and @data-element='header-bill-id']")
        print(bill_id.text)
    except:
        wait = WebDriverWait(driver, 10)
        bill_id = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='text-color-default' and @data-element='header-bill-id']")))
        print(bill_id.text)
    pt_name=driver.find_element(By.XPATH,'//div[@class="item box-padding-Bn box-fixed"]/h4[@data-element="patient-name"]')
    print(pt_name.text)
    pt_dob= driver.find_element(By.XPATH,'//div[@data-element="patient-dob"]/span[@class="text-color-default"]')
    print(pt_dob.text)
    pt_gender = driver.find_element(By.XPATH,'//div[@data-element="patient-gender"]/span[@class="text-color-default"]')
    print(pt_gender.text)
    pt_mobile= driver.find_element(By.XPATH,'//span[@data-element="patient-mobile-phone"]') 
    print(pt_mobile.text)
    if pt_mobile.text.strip()=='--': 
        pt_mobile= driver.find_element(By.XPATH,'//span[@data-element="patient-home-phone"]')
        print(pt_mobile.text, 'Home phone number found')
    pt_add= driver.find_element(By.XPATH,'//span[@data-element="patient-address"]')
    print(pt_add.text)
    time.sleep(2)
    ins_tab= driver.find_element(By.XPATH,'//li[@data-element="tab-insurance-details"]/div/button')
    driver.execute_script("arguments[0].scrollIntoView(true);", ins_tab)
    ins_tab.click()
    ins_name= driver.find_element(By.XPATH,'//div[@data-element="insurance-item-0"]//h4[@data-element="insurance-payer-name"]')
    ins_num= driver.find_element(By.XPATH,'//div[@data-element="insurance-item-0"]//div[@data-element="insurance-id"]/span[@class="text-color-default"]')
    print(ins_name.text)
    print(ins_num.text)

    try:
        service_details_tab= driver.find_element(By.XPATH,'//button[contains(., "Service details") and @type="button"]')
        service_details_tab.click()
    except:
        print('service details Tab is hidden')
    provider_name = driver.find_element(By.XPATH,'//div[@data-element="expandable-service-details"]//div[label[text()="Rendering Provider"]]/following-sibling::div/span[@data-element="performing-provider-name"]')
    place_of_service= driver.find_element(By.XPATH,'//div[@data-element="expandable-service-details"]//div[label[text()="Place of Service"]]/following-sibling::div/span[@data-element="place-of-service"]')
    facility_name= driver.find_element(By.XPATH,'//div[@data-element="expandable-service-details"]//div[label[text()="Facility"]]/following-sibling::div/span[@data-element="superbill-facility"]')
    print(provider_name.text, place_of_service.text, facility_name.text)
    services_data_list=[]
    for num in range(1,5):
        service_data=get_service_data(num)
        services_data_list.append(service_data)
    return (
        bill_id.text.strip(),pt_name.text.strip(),pt_dob.text.strip(),pt_gender.text.strip(),pt_mobile.text.strip(),pt_add.text.strip(),
        ins_name.text.strip(),ins_num.text.strip(),provider_name.text.strip(), place_of_service.text.strip(), 
        facility_name.text.strip(),services_data_list
            )
            

def create_daywise_temp():
    source_folder = r"C:\Automation\Python Automation\500 Charge Entry BOT"
    destination_folder = r"C:\Automation\Python Automation\500 Charge Entry BOT\BOT Status"
    file_to_copy = "Charge Entry Template.xlsx" 
    source_path = os.path.join(source_folder, file_to_copy)
    destination_path = os.path.join(destination_folder, f'{today_date}.xlsx')
    shutil.copy(source_path, destination_path)

def append_row_to_excel(file_path, data):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook['Sheet1']
    sheet.append(data)
    workbook.save(file_path)



def main_practice_fusion():
    create_daywise_temp() #creates day wise template for bot
    # login_practice_fusion(uid, pwd) #login to portal and waits untill user types security code
    start_driver()
    driver.maximize_window()
    show_bills_windows() #helps us to display bill rows 
    bill_links=get_bill_links() #collects all bill links i.e URLs in a list 
    count=1
    for url in bill_links:
        driver.get(url)
        pt_data = get_pt_details()
        print(pt_data)
        service1_cpt=pt_data[-1][0][0] if pt_data[-1][0]!=None else None
        service1_icds=pt_data[-1][0][1] if pt_data[-1][0]!=None else None
        service1_start_date=pt_data[-1][0][2] if pt_data[-1][0]!=None else None
        service1_end_date=pt_data[-1][0][3] if pt_data[-1][0]!=None else None
        service1_list=[service1_cpt, service1_icds, service1_start_date, service1_end_date]

        service2_cpt=pt_data[-1][1][0] if pt_data[-1][1]!=None else None
        service2_icds=pt_data[-1][1][1] if pt_data[-1][1]!=None else None
        service2_start_date=pt_data[-1][1][2] if pt_data[-1][1]!=None else None
        service2_end_date=pt_data[-1][1][3] if pt_data[-1][1]!=None else None
        service2_list=[service2_cpt, service2_icds, service2_start_date, service2_end_date]

        service3_cpt=pt_data[-1][2][0] if pt_data[-1][2]!=None else None
        service3_icds=pt_data[-1][2][1] if pt_data[-1][2]!=None else None
        service3_start_date=pt_data[-1][2][2] if pt_data[-1][2]!=None else None
        service3_end_date=pt_data[-1][2][3] if pt_data[-1][2]!=None else None
        service3_list=[service3_cpt, service3_icds, service3_start_date, service3_end_date]

        service4_cpt=pt_data[-1][3][0] if pt_data[-1][3]!=None else None
        service4_icds=pt_data[-1][3][1] if pt_data[-1][3]!=None else None
        service4_start_date=pt_data[-1][3][2] if pt_data[-1][3]!=None else None
        service4_end_date=pt_data[-1][3][3] if pt_data[-1][3]!=None else None
        service4_list=[service4_cpt, service4_icds, service4_start_date, service4_end_date]
        #('Robert Glass', '11/13/1939', 'Male', '(440) 781-1455', '21186 Lake Road, C/o Colleen Glass, Rocky River, OH 44116', 'Aetna 60054', '101160219600', 'Ashley Tompkins', '13 - Assisted Living Facility', 'Vitalia West Lake', [('99349', 'I48.0,G40.802,K59.81,M62.81', '02/09/2024', '02/09/2024'), None, None, None])
        pt_data= (count,)+pt_data[:-1] + tuple(service1_list) +tuple(service2_list) +tuple(service3_list) +tuple(service4_list)
        #update the data into excel below
        file_path=os.path.join(r"C:\Automation\Python Automation\500 Charge Entry BOT\BOT Status", f'{today_date}.xlsx')
        append_row_to_excel(file_path, pt_data)
        count+=1
        
    logout_btn=driver.find_element(By.XPATH,"//span[text()='Log out']")
    logout_btn.click()
    driver.close()
    driver.quit()

main_practice_fusion()







