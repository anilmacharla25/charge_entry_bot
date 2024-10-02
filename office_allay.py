#login to portal
#add patient to database if pt is not available in db
#add visit 
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
import time, openpyxl, pymsgbox
from pywinauto.keyboard import send_keys
import sys
import pyautogui

from fake_useragent import UserAgent
import os
os.system("cls")


def User():    
    ua = UserAgent()
    return ua.random

#edge options
edgeOption = webdriver.EdgeOptions()
edgeOption.add_argument(f'user-agent={User()}')
# edgeOption.add_experimental_option('excludeSwitches', ['enable-logging'])
# edgeOption.add_argument("start-maximized")
# edgeOption.add_argument("--enable-chrome-browser-cloud-management")
# Disable notifications
edgeOption.add_argument("--disable-notifications")
edgeOption.add_argument("--disable-popup-blocking")
edgeOption.add_argument("--disable-features=msEdgeEnableNurturingFramework")
edgeOption.add_argument('--disable-features=EdgeIdentityFeatures')


user_name_txt= 'XXXXXX'
password_txt= 'XXXXXXXXXX'

tday = datetime.datetime.today()
tdate = tday.strftime('%m-%d-%Y')


# Excel
# file_path = r"C:\Automation\Python Automation\500 Charge Entry BOT\BOT Status\02-13-2024.xlsx"
file_path = f"C:\\Automation\\Python Automation\\500 Charge Entry BOT\\BOT Status\\{tdate}.xlsx"
wbook = openpyxl.load_workbook(file_path)
sheet = wbook.active


driver = webdriver.Edge(options=edgeOption)
action = ActionChains(driver)
driver.implicitly_wait(10)
website_url = r"https://pm.officeally.com/pm/Login.aspx"
driver.get(website_url)
driver.maximize_window()
user_name= driver.find_element(By.ID,'username')
user_name.send_keys(user_name_txt)
pwd= driver.find_element(By.ID, 'password')
pwd.send_keys(password_txt)
continue_btn= driver.find_element(By.XPATH, "/html/body/main/section/div/div/div/form/div[2]/button")
continue_btn.click()
time.sleep(5)

try:
    driver.find_element(By.XPATH, '//button[contains(text(), "Close") and contains(@id, "pendo-button")]').click()
except:
    pass


def switch_window(window_no: int):
    handles=driver.window_handles
    driver.switch_to.window(handles[window_no])

def close_extra_windows():
    handles=driver.window_handles
    c=1
    while True:
        try:
            if len(handles)>1:
                driver.switch_to.window(handles[c])
                driver.close()
            else:
                break
        except:
            break
        c+=1


def search_pt(name, dob):
    search_box= driver.find_element(By.XPATH, "//input[@name='ctl00$phFolderContent$ucSearch$txtSearch']")
    search_box.clear()
    search_box.send_keys(name.split()[-1])
    time.sleep(2)
    driver.find_element(By.ID,'ctl00_phFolderContent_ucSearch_btnSearch').click()
    time.sleep(3)
    try:
        table= driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_myCustomGrid_myGrid"]')
        rows=table.find_elements(By.TAG_NAME,"tr")
        print(len(rows))
        for x in range(1, len(rows)):
            row_data = [cell.text.strip() for cell in rows[x].find_elements(By.TAG_NAME,'td')]
            print(row_data[8])
            if row_data[8] == dob:
                print('patient found in DB')
                return True
            else:
                print('Need to add patient')
                return False
    except:
        print('Need to add patient')
        return False


def add_pt_insurance_info(ins, group=None, subscriber=None):
    driver.find_element(By.XPATH, '//a[contains(text(), "Insurance")]').click()
    if group is not None:
        driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_InsuranceGroupNo"]').send_keys(group)
    if subscriber is not None:
        driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_InsuranceSubscriberID"]').send_keys(subscriber)
    driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_btnPopup"]').click()
    switch_window(1)
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_txtSearch"]').send_keys(ins.split()[0], Keys.ENTER)
    ins_table = driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_grvPopup"]/tbody')
    ins_rows = ins_table.find_elements(By.TAG_NAME, 'tr')
    if len(ins_rows) == 2:
        driver.find_element(By.XPATH, '//a[contains(text(), "Select")]').click()
    else:
        # pymsgbox.alert("Check!")
        pass

    switch_window(0)


def add_new_pt(name, gen, dob, addr, number,ins, group=None, subscriber=None):
    driver.find_element(By.XPATH, '//*[@id="addNewPatient"]').click()
    last_name_inp = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_LastName"]')
    last_name_inp.send_keys(name.split()[-1])
    print(name.split()[-1])
    first_name_inp = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_FirstName"]')
    first_name_inp.send_keys(name.split()[0])
    gender_inp = Select(driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_lstGender"]'))
    if gen.lower() == "female":
        gender_inp.select_by_visible_text("Female")
    elif gen.lower() == "male":
        gender_inp.select_by_visible_text("Male")
    else:
        gender_inp.select_by_visible_text("Unknown")
    print(dob)
    dob1 = dob.split('/')
    dob_m = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_DOB_Month"]')
    dob_m.send_keys(dob1[0])
    dob_d = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_DOB_Day"]')
    dob_d.send_keys(dob1[1])
    dob_y = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_DOB_Year"]')
    dob_y.send_keys(dob1[2])
    address_1 = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_AddressLine1"]')
    address_2 = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_AddressLine2"]')
    print(len(addr))
    if len(addr) == 4:
        address_1.send_keys(addr[0].strip())
        address_2.send_keys(addr[1].strip())
    elif len(addr) == 3:
        address_1.send_keys(addr[0].strip())
    city = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_City"]')
    city.send_keys(addr[-2].strip())
    state = Select(driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_lstState"]'))
    state.select_by_visible_text(addr[-1].split()[0])
    pin = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_Zip"]')
    pin.send_keys(addr[-1].split()[1])
    if number is not None:
        mobile_inp1 = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_CellPhone_AreaCode"]')
        mobile_inp1.send_keys(number[1:4])
        mobile_inp2 = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_CellPhone_Prefix"]')
        mobile_inp2.send_keys(number[6:9])
        mobile_inp3 = driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_CellPhone_Number"]')
        action.move_to_element(mobile_inp3).click().send_keys(number[10:]).perform()
    add_pt_insurance_info(ins, group, subscriber)
    # pymsgbox.alert("Check!")
    # driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_btnCancel"]').click()       # Cancel button
    driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucPatient_btnUpdate"]').click()       # add patient button


def billing_options(facility):
    driver.find_element(By.XPATH, '//a[contains(text(), "Billing Options")]').click()
    driver.find_element(By.XPATH, '//*[@title="Facility List"]').click()
    switch_window(1)
    driver.get(r'https://pm.officeally.com/pm/SharedFiles/popup/Popup.aspx?name=Facilities')
    time.sleep(2)
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Facility Name")]').click()
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlCondition"]/option[contains(text(), "Starts With")]').click()
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_txtSearch"]').send_keys(facility.split()[0], Keys.ENTER)
    facility_table = driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_grvPopup"]/tbody')
    facility_rows = facility_table.find_elements(By.TAG_NAME, 'tr')
    for x in range(1, len(facility_rows)):
        try:
            facility_row_data = [cell.text.strip() for cell in facility_rows[x].find_elements(By.TAG_NAME,'td')]
            if facility.split()[0] in facility_row_data[1]:
                facility_rows[x].find_element(By.XPATH, '//td/a[contains(text(), "Select")]').click()
        except:
            try:
                driver.close()
            except:
                pass
            print('facility row not found')
        
    switch_window(0)


def no_of_units(icds: list):
    count = 0
    units = ""
    for unit in range(0, len(icds)):
        count+=1
        if count==5:
            break
        units = units + str(count)
    return units


def billing_info(icds: list, cpts:list, start_dates: list, end_dates: list):
    #to close unneccesary windows which were already opened
    close_extra_windows()
    driver.switch_to.window(driver.window_handles[0])            
    driver.find_element(By.XPATH, '//a[contains(text(), "Billing Info")]').click()
    driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_ucDiagnosisCodes_rdICD_10"]').click()
    icd_boxes = driver.find_elements(By.XPATH, '//*[@class="textbox dc dc_10 ui-autocomplete-input"]')
    unit_boxes = driver.find_elements(By.XPATH, '//*[@class="textbox cptPointer js-change"]') #
    time.sleep(4)
    for x in range(0, len(icds)):
        print(icd_boxes[x])
        print(icds[x])
        time.sleep(5)
        print('the index value',x)
        print('the length of the icds boxes', len(icd_boxes))
        try:
            icd_boxes[x].send_keys(icds[x])
        except Exception as e:
            print('the icd error', e)
            print('icd not poping up')
            break

        time.sleep(2)
        send_keys('{DOWN}')
        time.sleep(1)
        send_keys('{ENTER}')
        print(icds[x])
    close_extra_windows()
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    cpt_boxes = driver.find_elements(By.XPATH, "//td[@class='LIBorderRight' and @align='center']/input[@type='button' and @class='button' and @title='User Procedure Codes']")
                                                #//input[@title="User Procedure Codes"]
    if len(cpt_boxes)==0:
        print('CPT boxes are Not Found')
        pyautogui.scroll(200)
        # element_to_scroll_to= driver.find_element(By.XPATH, "//input[@type='button' and @class='button' and @onclick='popupSuperbill()']")
        # driver.execute_script("arguments[0].scrollIntoView(true);", element_to_scroll_to)
        cpt_boxes = driver.find_elements(By.XPATH, "//td[@class='LIBorderRight' and @align='center']/input[@type='button' and @class='button' and @title='User Procedure Codes']")
           
    start_date_boxes = driver.find_elements(By.XPATH, '//*[@class="textbox cptDOSFrom js-change"]')
    end_date_boxes = driver.find_elements(By.XPATH, '//*[@class="textbox cptDOSTo js-change"]')
    for x in range(0, len(cpts)):
        print('the length of cpt boxes', len(cpt_boxes))
        print('the index value', x)
        if len(cpt_boxes)==0:
            driver.refresh()
            break
        cpt_boxes[x].click()
        handles=driver.window_handles
    
        driver.switch_to.window(handles[1])
        driver.get('https://pm.officeally.com/pm/SharedFiles/popup/Popup.aspx?name=UserCPT')
        driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_txtSearch"]').send_keys(cpts[x], Keys.ENTER)
        time.sleep(2)
        driver.find_element(By.XPATH, '//a[contains(text(), "Select")]').click()

        driver.switch_to.window(handles[0])
        time.sleep(2)
        start_date_boxes[x].clear()
        start_date_boxes[x].send_keys(start_dates[x])
        end_date_boxes[x].send_keys(end_dates[x])
        unit_boxes[x].clear()
        unit_boxes[x].send_keys(no_of_units(icds=icds))



def add_new_visit(name, dob, prov_name, facility, icds: list, cpts: list, start_dates:list, end_dates:list):
    driver.find_element(By.XPATH, '//*[@id="addNewVisit"]').click()
    # Select patient
    driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_Button1"]').click()
    time.sleep(2)
    try:
        switch_window(1)
        # driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]').click()
        l_name = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]')))
        l_name.click()
    except:
        try:
            switch_window(1)
            driver.get('https://pm.officeally.com/pm/SharedFiles/popup/Popup.aspx?name=Patient&returnData=fullpatient')
            # driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]').click()
            l_name = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]')))
            l_name.click()
        except:
            driver.close()
            switch_window(1)
            l_name = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]')))
            l_name.click()


    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlCondition"]/option[contains(text(), "Starts With")]').click()
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_txtSearch"]').send_keys(name.split()[-1], Keys.ENTER)
    time.sleep(2)
    pt_table = driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_grvPopup"]/tbody')
    pt_rows = pt_table.find_elements(By.TAG_NAME, 'tr')
    # print(pt_rows)
    for r in pt_rows:
        print([cell.text.strip() for cell in r.find_elements(By.TAG_NAME,'td')])
    # time.sleep(10)
    # sys.exit()
    for x in range(1, len(pt_rows)):
        pt_row_data = [cell.text.strip() for cell in pt_rows[x].find_elements(By.TAG_NAME,'td')]
        print(pt_row_data)
        print(dob)
        print(pt_row_data[7])
        # sys.exit()
        if pt_row_data[7] == dob:
            pt_rows[x].find_element(By.XPATH, '//td/a[contains(text(), "Select")]').click()
            break
    switch_window(0)
    # Select provider
    driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_Button2"]').click()
    try:
        switch_window(1)
        driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]').click()
    except:
        driver.get('https://pm.officeally.com/pm/SharedFiles/popup/Popup.aspx?name=Provider')
        driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlSearch"]/option[contains(text(), "Last Name")]').click()
        

    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_ddlCondition"]/option[contains(text(), "Starts With")]').click()
    if len(prov_name.split())>2:
        provider_last_name=prov_name.split()[-2] +' '+ prov_name.split()[-1] 
    else:
        provider_last_name=prov_name.split()[-1] 
    driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_txtSearch"]').send_keys(provider_last_name, Keys.ENTER)
    time.sleep(2)
    prov_table = driver.find_element(By.XPATH, '//*[@id="ctl04_popupBase_grvPopup"]/tbody')
    prov_rows = prov_table.find_elements(By.TAG_NAME, 'tr')
    for x in range(1, len(prov_rows)):
        prov_row_data = [cell.text.strip() for cell in prov_rows[x].find_elements(By.TAG_NAME,'td')]
        print(prov_row_data[2])
        if prov_name.split()[0] in prov_row_data[2]:
            prov_rows[x].find_element(By.XPATH, '//td/a[contains(text(), "Select")]').click()
    switch_window(0)
    billing_options(facility)
    try:
        billing_info(icds, cpts, start_dates, end_dates)
        # pymsgbox.alert("Check!")
        driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_btnCancel"]').click() #cancel button
        # driver.find_element(By.XPATH, '//*[@id="ctl00_phFolderContent_btnUpdate"]').click() #update button
        return 'Updated'
    except:
        print('Error Occured')
        return 'Error'


for row in sheet.iter_rows(min_row=2):
    print(row[30].value)
    if row[30].value =="Yes":
        continue
    pt_name = row[2].value
    pt_dob = row[3].value
    pt_gender = row[4].value
    pt_no = row[5].value
    pt_addr = row[6].value.split(',')
    ins_name = row[7].value
    ins_id = row[8].value
    provider = row[9].value
    facility_name = row[11].value
    try:
        icds = row[13].value.split(',')
        cpts = [row[12].value, row[16].value, row[20].value, row[24].value]
    except:
        icds= None
        cpts = None

    # try: #outer try block
    for x in range(4):
        try:
            cpts.remove(None)
        except:
            pass
    print(cpts)
    start_dates = [row[14].value, row[18].value, row[22].value, row[26].value]
    print(start_dates)
    end_dates = [row[15].value, row[19].value, row[23].value, row[27].value]
    subscriber_id = row[28].value
    group_id = row[29].value
    time.sleep(2)
    print('clicking on manage patients')
    driver.find_element(By.XPATH, '//span[contains(text(), "Manage Patients")]').click()
    pt_found = search_pt(name=pt_name, dob=pt_dob)
    if not pt_found:
        add_new_pt(name=pt_name, dob=pt_dob, gen=pt_gender, number=pt_no, addr=pt_addr, group=group_id, subscriber=subscriber_id, ins=ins_name)
        time.sleep(2)
        print('Patient Added!')
    
    # Patient Visits
    driver.find_element(By.XPATH, '//*[@id="patient-visits_tab"]/span').click()
    time.sleep(2)
    visit_update_status=add_new_visit(name=pt_name, dob=pt_dob, prov_name=provider, facility=facility_name, icds=icds, cpts=cpts, start_dates=start_dates, end_dates=end_dates)
    if visit_update_status=='Error':
        row[30].value = "No"
    elif visit_update_status=='Updated':
        row[30].value = "Yes"
    wbook.save(file_path)
    # except:
    #     row[31].value = "No"
    #     wbook.save(file_path)


driver.quit()
