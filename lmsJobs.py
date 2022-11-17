from asyncio.log import logger
import datetime as date
from time import sleep
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from os import path
import pandas as pd
import os
from selenium.webdriver.common.keys import Keys
class DriverBuilder():
    def setting_chrome_options(self):
        # chrome_options = Options()
        # chrome_options.add_argument('--headless')
        options = Options()
        options.headless = True
        options.add_argument("start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox')
        # options.binary_location('D:\lmsapi\chromedriver')
        # options.add_argument('--disable-infobars')
        # options.add_argument('--disable-extensions')
        # options.add_argument('--disable-gpu')
        return options
    def get_driver(self):
        try:
            service = Service(executable_path="chromedriver.exe")
            # driver = webdriver.Chrome(service=service,options = self.setting_chrome_options())
            # driver = webdriver.Chrome(service = ChromeService(ChromeDriverManager().install(),options=self.setting_chrome_options()))
            driver = webdriver.Chrome(service = service, options = self.setting_chrome_options())
            # chromedriver_autoinstaller.install()
            # driver = webdriver.Chrome(options = self.setting_chrome_options())
            return driver
        except Exception as e:
            logger.error('Failed to get driver: ' + str(e))
            print("Failed!!!!!")
class MovingCenterLMS():
    def moving_center_LMS(self,driver,student_info):
        center_dict_visang = {
            "QT" : "QT",
            "Wallace Franchise" : "Wallace Franchise",
            "eduplanning" : "eduplanning",
            "Math-TEST" : "Math-TEST",
            "VALAB": "V-ALAB",
            "PDL" : "PDL",
            "KDV" : "KDV",
            "HB" : "HBC",
            "CHG" : "CHC",
            "LTK" : "LTK",
            "PVT" : "PVT",
            "CSM" : "CSM",
            "ALAB TEST" : "ALAB TEST",
            "3cell_center": "3cell_center",
            "Visang_T003" : "Visang_T003",
            "Visang_T001" : "Visang_T001",
            "Visang_T002" : "Visang_T002",
            "본사관리가맹" : "본사관리가맹"
        }
        center_dict = {
            "QT" : "1759",
            "Wallace Franchise" : "1758",
            "eduplanning" : "1757",
            "Math-TEST" : "1756",
            "V-ALAB": "1755",
            "PDL" : "1754",
            "KDV" : "1753",
            "HBC" : "1752",
            "CHC" : "1751",
            "LTK" : "1750",
            "PVT" : "1749",
            "CSM" : "1748",
            "ALAB TEST" : "1747",
            "3cell_center": "1741",
            "Visang_T003" : "1739",
            "Visang_T001" : "1738",
            "Visang_T002" : "1737",
            "본사관리가맹" : "1"
        }
        student_id = student_info['MaHV']
        # Mapping center code between Flow and Visang
        center_in = center_dict_visang[student_info['CenterMoving']]
        center_out = center_dict_visang[student_info['CenterCurrent']]
        # head to List of Student Status page
        driver.get('https://mag.alab.edu.vn/member/memberSearchList.do')
        sleep(3)
        # Click on the drop-down to select ID field
        WebDriverWait(driver=driver,timeout=10).until(
            EC.element_to_be_clickable(
                (By.XPATH,'//*[@id="inStudentType"]/option[2]'))
            ).click()
        sleep(2)
        # Enter student_id to the search box
        WebDriverWait(driver=driver,timeout=10).until(
            # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
            EC.element_to_be_clickable(
                (By.ID,'inStudentSearch'))
            ).send_keys(student_id)
        sleep(3)
        # Click on the Search button
        WebDriverWait(driver=driver,timeout=10).until(
            # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR,"button[onclick^='searchData();']"))
            ).click()
        sleep(2)
        # Checking valid student ID with rows > 1
        student_rows = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr')
        if(len(student_rows) == 1):
            sleep(5)
            # Checking valid student ID with rows = 0
            current_center = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[3]').text
            current_name = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[4]/a').text
            if(len(current_center)):
                # Check the current student's center with the center out AND check student ID with student name
                # If matches, do the moving process
                # Else, reply mail via Power Automate Flow.
                if ((current_center == center_out) & (current_name == student_info['Name'])):
                    
                    # Click on Application button from the desired student
                    WebDriverWait(driver=driver,timeout=10).until(
                        # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[9]/a'))
                        ).click()
                    print("Application is clicked!!!!!!")
                    sleep(5)
                # Checking center_in is on the next page (Paging!!!!)
                    radio_buttons = driver.find_elements(By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                    
                    if(len(radio_buttons)):
                        radio_button = driver.find_element(By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                        driver.execute_script('arguments[0].scrollIntoView();',radio_button)
                        WebDriverWait(driver,20).until(
                            EC.element_to_be_clickable(
                                (By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                            )
                        ).click()
                        print("Found the center in the 1st page!!!!!")
                        sleep(3)
                    else:
                        print("No center found in the 1st page!!!!!!")
                        # Click on the next Page
                        WebDriverWait(driver=driver,timeout=10).until(
                        # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
                            EC.element_to_be_clickable(
                                    (By.XPATH,'//*[@id="choiceClass"]/div[1]/div/div/nav/ul/li[4]/a[contains(@onclick,"onPopChangePage(2); return false;")]'))
                            ).click()
                        print('Go to the next page!!!!!!')
                        sleep(2)
                        # Choose the center on the next page
                        radio_buttons = driver.find_elements(By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                        
                        if(len(radio_buttons)):
                            radio_button = driver.find_element(By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                            driver.execute_script('arguments[0].scrollIntoView();',radio_button)
                            WebDriverWait(driver,20).until(
                                EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR,'input#changeFridx[value="{}"]'.format(center_dict[center_in]))
                                )
                            ).click()
                            print("Found the center in the 2nd page!!!!!")
                        else:
                            print('element not found!!!!!')
                        sleep(2)

                    # Click on Save button on the pop-up page
                    WebDriverWait(driver,20).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR,"#choiceClass > div.modal-footer > button")
                        )
                    ).click()
                    sleep(1)
                    # Click on OK button on 1st popup window
                    driver.switch_to.alert.accept()
                    sleep(1)
                    # Click on OK button on 2nd popup window
                    driver.switch_to.alert.accept()
                    sleep(2)
                    # Click on the Search button
                    WebDriverWait(driver=driver,timeout=10).until(
                        # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR,"button[onclick^='searchData();']"))
                    ).click()
                    current_center = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[3]').text
                    if(current_center == center_in):
                        return {
                            'Name' : student_info['Name'],
                            'MaHV' : student_info['MaHV'],
                            'CenterCurrent' : current_center,
                            'Message' : "Moving Successful!!!!!"
                        }
                    else:
                        return{
                            'Name': student_info['Name'],
                            'MaHV' : student_info['MaHV'],
                            'CenterCurrent' : student_info['CenterCurrent'],
                            'Message': 'Moving Process Failed!!!!!!'
                        }
                else:
                    ### DO SOMETHING HERE (i.e, REPLY MAIL TO AO)
                    return {
                        'Name': student_info['Name'],
                        'MaHV': student_info['MaHV'],
                        'CenterCurrent': student_info['CenterCurrent'],
                        'Message' : "Not matching the current student's center OR ID and Student Name are not matching!!!!!"
                    }
            else:
                return {
                    'Name': student_info['Name'],
                    'MaHV': student_info['MaHV'],
                    'CenterCurrent': student_info['CenterCurrent'],
                    'Message' : "Invalid Student ID"
                    }
            
        else:
            return {
                    'Name': student_info['Name'],
                    'MaHV': student_info['MaHV'],
                    'CenterCurrent': student_info['CenterCurrent'],
                    'Message' : "Invalid Student ID"
                }
class LoginAdminLMS():
    def login_mag_admin(self,driver):
        # Mag Admin credentials
        username = "Alab"
        password = "AlabmAg2022"
        try:
            # head to Mag login page
            driver.get("https://mag.alab.edu.vn")
            # find ID field and send the username itself to the ID input field
            driver.find_element("id","userId").send_keys(username)
            # find Password field and send the password itself to the password input field
            driver.find_element("id","userPwd").send_keys(password)
            # click login button
            driver.find_element("id","btnLogin").click()
            # wait the ready state to be complete
            WebDriverWait(driver=driver, timeout=10).until(
                lambda x: x.execute_script("return document.readyState === 'complete'")
            )
            print("Singed-in Mag Admin")
        except Exception as e: 
            logger.error('Failed to login: '+ str(e))
            driver.quit()
    def login_center(self,driver,login_code,center_code):
        # self.login_mag_admin(driver)
        # head to center info page
        driver.get('https://mag.alab.edu.vn/branchfranchise/franchiseAccountList.do')
        sleep(2)
        # Find the center login button with Paging
        login_buttons = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code))        
        sleep(3)
        if(len(login_buttons)):
            # store the center list page as the first window
            admin_page = driver.current_window_handle    
            login_button = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code))
            driver.execute_script('arguments[0].scrollIntoView();',login_button)
            #Click to Login to Center
            WebDriverWait(driver=driver,timeout=10).until(
                EC.element_to_be_clickable(
                    (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code)))
            ).click()
            # Click on OK button on alert window to login
            driver.switch_to.alert.accept()
            sleep(2)
            print("Found the center in the 1st page!!!!!")
            print("Signed in to the center {}".format(center_code))
            # Handler to navigate to specific center page
            # after opening center page, change window handle
            for handle in driver.window_handles: 
                if handle != admin_page: 
                    center_page = handle
                    driver.switch_to.window(center_page)
                    # Now we are on Center Page
                    sleep(5)
                    # Store center page in window handler
                    center_page_window = driver.current_window_handle
                    return [admin_page, center_page_window]
        else:
            print("No center found in the 1st page!!!!!!")
            # Click on the next Page
            WebDriverWait(driver=driver,timeout=10).until(
                EC.element_to_be_clickable(
                    (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div[3]/div/nav/ul/li[4]/a[contains(@onclick,"linkPage(2); return false;")]'))
                ).click()
            print('Go to the next page!!!!!!')
            sleep(2)
            #Find the center on the 2nd page
            login_buttons = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code))
            if(len(login_buttons)):
                # store the center list page as the first window
                admin_page = driver.current_window_handle
                login_button = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code))
                driver.execute_script('arguments[0].scrollIntoView();',login_button)
                #Click to Login to Center
                WebDriverWait(driver=driver,timeout=10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[11]/a[@onclick ="frLogin(\'{}\')"]'.format(login_code)))
                ).click()
                # Click on OK button on alert window to login
                driver.switch_to.alert.accept()
                sleep(2)
                print("Found the center in the 2nd page!!!!!")
                print("Signed in to the center {}".format(center_code))
                # Handler to navigate to specific center page
                # after opening center page, change window handle
                for handle in driver.window_handles: 
                    if handle != admin_page: 
                        center_page = handle
                        driver.switch_to.window(center_page)
                        # Now we are on Center Page
                        sleep(5)
                        # Store center page in window handler
                        center_page_window = driver.current_window_handle
                        return [admin_page, center_page_window]
class download_LMS():
    def enable_download(self,driver):
        download_dir = path.dirname(path.realpath(__file__))
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_dir}}
        driver.execute("send_command", params)

    def download_branch_info_file(self,driver):
        # head to branch center info page
        driver.get("https://mag.alab.edu.vn/branchfranchise/branchAccountList.do")
        # click on the branch-specific info button 
        WebDriverWait(driver=driver,timeout=10).until(
            # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR,"button[onclick^='brListExcel2();']"))
            ).click()
    def download_total_student_file(self,driver):
        # head to branch center info page
        driver.get("https://mag.alab.edu.vn/member/adminChargeTargetAttendList.do")
        WebDriverWait(driver=driver,timeout=10).until(
            # driver.find_element(By.CSS_SELECTOR,"button[onclick^='brListExcel2();']").click())
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR,"button[onclick^='listExcel(0);']"))
            ).click()
    def upload_member_template(self, driver):
        print('Uploading member template')
class ResetPassword():
    def read_excel_file(self,file_name):
        teacher_df = pd.read_excel(file_name,sheet_name='Teacher List')
        return teacher_df
    def reset_pwd_AO(self,driver,password):
        # center_dict = {
        #     "QT" : "1759",
        #     "V-ALAB": "1755",
        #     "PDL" : "1754",
        #     "KDV" : "1753",
        #     "HBC" : "1752",
        #     "CHC" : "1751",
        #     "LTK" : "1750",
        #     "PVT" : "1749",
        #     "CSM" : "1748"
        #     # "ALAB TEST" : ["1747","8Rm2Ute2MsRHyuvvzsnW3w"]
        # }
        center_dict = {
            "ALAB TEST" : "1747"
            # "QT" : "1759"
        }
        #### Login to Mag Admin ######
        login = LoginAdminLMS()
        login.login_mag_admin(driver)
        # head to center info page
        driver.get('https://mag.alab.edu.vn/branchfranchise/franchiseAccountList.do')
        sleep(2)
        for center_code in center_dict:
            password_by_center = self.pwd_gen(center_code,password)
            # Find the center with Paging
            centers = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code]))        
            sleep(3)

            if(len(centers)):
                center = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code]))
                driver.execute_script('arguments[0].scrollIntoView();',center)
                #Click to Center Account
                WebDriverWait(driver=driver,timeout=10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code])))
                ).click()
                print("Found the center in the 1st page!!!!!")
                sleep(3)
                pwd_buttons = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]')
                if(len(pwd_buttons)):
                    pwd_button = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]')
                    driver.execute_script('arguments[0].scrollIntoView();',pwd_button)
                    # before clicking button to open popup, store the current window handle
                    main_window = driver.current_window_handle
                    #Click on the Change PWD button
                    WebDriverWait(driver=driver,timeout=10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]'))
                    ).click()
                    # after opening popup, change window handle
                    for handle in driver.window_handles: 
                        if handle != main_window: 
                            popup = handle
                            driver.switch_to.window(popup)
                    # Now should be on popup window
                    # Then, fill new pwd
                    sleep(2)
                    driver.find_element("id","pwd1").send_keys(password_by_center)
                    sleep(2)
                    # Fill confirm pwd
                    driver.find_element("id","pwd2").send_keys(password_by_center)
                    # Click on Save button
                    WebDriverWait(driver=driver,timeout=10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,'//*[@id="pop_footer"]/a[@onclick = "frmSubmit()"]'))
                    ).click()
                    sleep(2)
                    # Click on Save button on 1st alert window
                    driver.switch_to.alert.accept()
                    sleep(1)
                    # Click on OK button on 2nd alert window
                    driver.switch_to.alert.accept()
                    driver.switch_to.window(main_window)
                    print('Password has been changed in center: {}'.format(center_code))
                    print('New Password: {}'.format(password_by_center))
                else: 
                    print('Not find Change PWD button!!!!!')
                # Head back to Center Info Page
                driver.get('https://mag.alab.edu.vn/branchfranchise/franchiseAccountList.do')
                sleep(2)
            else:
                print("No center found in the 1st page!!!!!!")
                # Click on the next Page
                WebDriverWait(driver=driver,timeout=10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div[3]/div/nav/ul/li[4]/a[contains(@onclick,"linkPage(2); return false;")]'))
                    ).click()
                print('Go to the next page!!!!!!')
                sleep(2)
                #Find the center on the 2nd page
                centers = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code]))
                if(len(centers)):
                    center = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code]))
                    driver.execute_script('arguments[0].scrollIntoView();',center)
                    #Click to Center Account on the 2nd page
                    WebDriverWait(driver=driver,timeout=10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[2]/a[@onclick = "frMod(\'{}\')"]'.format(center_dict[center_code])))
                    ).click()
                    print("Found the center in the 2nd page!!!!!")
                    sleep(3)
                    pwd_buttons = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]')
                    if(len(pwd_buttons)):
                        pwd_button = driver.find_element(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]')
                        driver.execute_script('arguments[0].scrollIntoView();',pwd_button)
                        # before clicking button to open popup, store the current window handle
                        main_window = driver.current_window_handle
                        #Click on the Change PWD button
                        WebDriverWait(driver=driver,timeout=10).until(
                            EC.element_to_be_clickable(
                                (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div/button[@onclick = "fnFrUserPwdPop()"]'))
                        ).click()
                        # after opening popup, change window handle
                        for handle in driver.window_handles: 
                            if handle != main_window: 
                                popup = handle
                                driver.switch_to.window(popup)
                        # Now should be on the popup window
                        # Then, fill new pwd
                        sleep(2)
                        driver.find_element("id","pwd1").send_keys(password_by_center)
                        sleep(2)
                        # Fill confirm pwd
                        driver.find_element("id","pwd2").send_keys(password_by_center)
                        # Click on Save button
                        WebDriverWait(driver=driver,timeout=10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="pop_footer"]/a[@onclick = "frmSubmit()"]'))
                        ).click()
                        sleep(2)
                        # Click on Save button on 1st alert window
                        driver.switch_to.alert.accept()
                        sleep(1)
                        # Click on OK button on 2nd alert window
                        driver.switch_to.alert.accept()
                        driver.switch_to.window(main_window)
                        print('Password has been changed in center: {}'.format(center_code))
                        print('New Password: {}'.format(password_by_center))
                    else: 
                        print('Not find Change PWD button!!!!!')
                    # Head back to Center Info Page
                    driver.get('https://mag.alab.edu.vn/branchfranchise/franchiseAccountList.do')
                    sleep(2)
                else:
                    print('No center found!!!!!') 
        # Close driver
        driver.quit()   
    def reset_pwd_teachers(self,driver,password):
        #### Read Teacher List from excel file
        teacher_df = self.read_excel_file('Teacher-List.xlsx')
        # center_dict = {
        #     "QT" : ["1759","krEJcwgxb2nY5_0DNHmEeQ"],
        #     "V-ALAB": ["1755","JOvs6TBc8wbJqBY1_z6lQQ"],
        #     "PDL" : ["1754","L6pbhxGqfHfq0BBsSfy4tQ"],
        #     "KDV" : ["1753","gpl6yD6RxUKR02pUCzf_7g"],
        #     "HBC" : ["1752","aOoeVoHC2ZJhq6LtXy0HVQ"],
        #     "CHC" : ["1751","EtOfGnY2FhzePPL63VzvvQ"],
        #     "LTK" : ["1750","picFaVF3cmwbeTOJI7QFMA"],
        #     "PVT" : ["1749","vceFOCTEvpUWhbZZS3l8sA"],
        #     "CSM" : ["1748","BQ_8-DgQA0SkZSbtQkNO-A"],
        #     "ALAB TEST" : ["1747","8Rm2Ute2MsRHyuvvzsnW3w"]
        # }
        center_dict = {
            "ALAB TEST" : ["1747","8Rm2Ute2MsRHyuvvzsnW3w"]
            # "QT" : ["1759","krEJcwgxb2nY5_0DNHmEeQ"]
        }
        #### Login to Mag Admin ######
        login = LoginAdminLMS()
        login.login_mag_admin(driver)
        for center_code in center_dict:
            password_by_center = self.pwd_gen(center_code,password)
            login_code = center_dict[center_code][1]
            windows = login.login_center(driver,login_code,center_code)
            admin_page = windows[0]
            center_page_window = windows[1]
            teacher_by_center_df = teacher_df.loc[teacher_df["Center's name"] == center_code].reset_index()
            sleep(5)
            # Head to Teacher Management
            driver.get('https://join.alab.edu.vn/franchisemanage/lectureList.do')
            # Store center page in window handler
            # center_page_window = driver.current_window_handle
            # Choose ID Field
            WebDriverWait(driver=driver,timeout=10).until(
                EC.element_to_be_clickable(
                    (By.XPATH,'//*[@id="frmMain"]/div[1]/div[2]/div[2]/select/option[2]'))
            ).click()
            sleep(3)
            for i in range(len(teacher_by_center_df)):
                teacher_id = teacher_by_center_df.loc[i]['Teacher ID']
                sleep(2)
                # Clear teacher_id in search box if needed
                driver.find_element("id","searchText").clear()
                sleep(2)
                # Fill teacher_id to search box
                driver.find_element("id","searchText").send_keys(teacher_id)
                sleep(2)
                # Click on search button
                WebDriverWait(driver=driver,timeout=10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH,'//*[@id="buttonSearch" and @onclick="searchLecturerData()"]'))
                ).click()
                sleep(3)
                # Click on the searched teacher info
                WebDriverWait(driver=driver,timeout=10).until(
                EC.element_to_be_clickable(
                    (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td[3]/a'))
                ).click()
                sleep(2)
                # Scroll into pwd button
                pwd_buttons = driver.find_elements(By.XPATH,'//*[@id="frmMain"]/div/button[2]')
                if(len(pwd_buttons)):
                    pwd_button = driver.find_element(By.XPATH,'//*[@id="frmMain"]/div/button[2]')
                    driver.execute_script('arguments[0].scrollIntoView();',pwd_button)
                    # Click on pwd button
                    WebDriverWait(driver=driver,timeout=10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="frmMain"]/div/button[2]'))
                    ).click()
                    # Switching to popup window for changing pwd
                    for handle in driver.window_handles:
                        if ((handle != admin_page) & (handle != center_page_window)):
                            popup_pwd = handle
                            driver.switch_to.window(popup_pwd)
                            # NOW WE ARE IN POP-UP WINDOW
                            # Checking teacher id with default teacher id
                            default_id = driver.find_element(By.XPATH,'//*[@id="frmMain"]/div[1]/table/tbody/tr[3]/td/div/div/input').get_attribute('value')
                            if (default_id == teacher_id):
                                # Enter the password
                                driver.find_element(By.XPATH,'//*[@id="pwd1"]').send_keys(password_by_center)
                                sleep(2)
                                # Enter the confirm password
                                driver.find_element(By.XPATH,'//*[@id="pwd2"]').send_keys(password_by_center)
                                sleep(3)
                                #Click on Edit button to save password
                                WebDriverWait(driver=driver,timeout=10).until(
                                    EC.element_to_be_clickable(
                                        (By.XPATH,'//*[@id="submit2" and @onclick="modifyPwd();"]'))
                                ).click()
                                sleep(2)
                                # Click OK on 1st alert
                                driver.switch_to.alert.accept()
                                sleep(2)
                                # Click OK on 2nd alert
                                driver.switch_to.alert.accept()
                                sleep(2)
                                # Switch back to center page
                                driver.switch_to.window(center_page_window)
                                # Click on Previous button
                                WebDriverWait(driver=driver,timeout=10).until(
                                    EC.element_to_be_clickable(
                                        (By.XPATH,'//*[@id="frmMain"]/div/button[@onclick ="goConfirmList()"]'))
                                ).click()
                                sleep(2)
                                # Click on alert to confirm
                                driver.switch_to.alert.accept()
                                print('Password has been changed in center: {}'.format(center_code))
                                print('New Password: {}'.format(password_by_center))
                            else:
                                print('ID is not matching!!!!! Check again!!!!!!')
                        else:
                            print("Not the right window!!!!!")
                else:
                    print('Password button is not found!!!!!!')
            # Switch back to admin page
            driver.switch_to.window(admin_page)
            sleep(2)
            # Head to Center List page in admin page
            driver.get('https://mag.alab.edu.vn/branchfranchise/franchiseAccountList.do') 
        # Close the driver
        driver.quit()   
    def pwd_gen(self,center_code, password):
        password_by_center = ""
        if(center_code == "V-ALAB"):
            # Add center code to the end of password
            password_by_center = "{}{}".format(password,"ONL")
        elif(center_code == "ALAB TEST"):
            password_by_center = "{}{}".format(password,"TEST")
        else:
            password_by_center = "{}{}".format(password,center_code)
        return password_by_center
class CreateStudent():
    def create_students(self, driver,student_info):
        # center_dict = {
        #     "QT" : ["1759","krEJcwgxb2nY5_0DNHmEeQ"],
        #     "V-ALAB": ["1755","JOvs6TBc8wbJqBY1_z6lQQ"],
        #     "PDL" : ["1754","L6pbhxGqfHfq0BBsSfy4tQ"],
        #     "KDV" : ["1753","gpl6yD6RxUKR02pUCzf_7g"],
        #     "HBC" : ["1752","aOoeVoHC2ZJhq6LtXy0HVQ"],
        #     "CHC" : ["1751","EtOfGnY2FhzePPL63VzvvQ"],
        #     "LTK" : ["1750","picFaVF3cmwbeTOJI7QFMA"],
        #     "PVT" : ["1749","vceFOCTEvpUWhbZZS3l8sA"],
        #     "CSM" : ["1748","BQ_8-DgQA0SkZSbtQkNO-A"],
        #     "ALAB TEST" : ["1747","8Rm2Ute2MsRHyuvvzsnW3w"]
        # }
        center_dict = {
            "ALAB TEST" : ["1747","8Rm2Ute2MsRHyuvvzsnW3w","3119",'ALAB-TEST'],
            "QT" : ["1759","krEJcwgxb2nY5_0DNHmEeQ","3120","Quang Trung"]
        }
        #### MESSAGE DICTIONARY TO HOLD THE MESSAGE WITH ERROR NUMBER #####
        message_dict = {
            #student id existed
            "student information needs to be corrected for a total of 1 enrollments.  Do you want to proceed?" : [1,"Student ID Existed"],
            #student id empty
            "Please enter ID.":[2, "Student ID empty"],
            #student password empty
            "Please enter PW.":[3, "Student Password empty"],
            #student name empty
            "Please enter Name." : [4,"Student Name empty"],
            #Class code empty
            "Class code is mandatory to be filled." :[5, "Class Code empty"],
            #wrong class code
            "Class code is not a class belonging to the institution." : [6, "Wrong Class Code"],
            #Parent name empty
            "Please enter the parent’s names.": [7, "Parent Name empty"],
            #Parent code empty
            "Please enter parent classification.": [8, "Parent Code empty"],
            #Parent code with string format
            "Correction and batch registration failed.   Please contact your system representative.": [9, "Parent Code with wrong format"],
            #Parent Passwod is empty
            "parentsPlease enter the password.":[10, "Parent Password empty"],
            #Create new student (Success meassage)
            "1 modifications and registrations have been completed." : [200, "Create New Student Successfully"]
        }
        # Login Mag Admin
        login = LoginAdminLMS()
        login.login_mag_admin(driver)
        for student in student_info:
            #### DEFINE VARIABLES ######
            ems_id = student["ems_id"]
            # Mapping Center Code between EMS and LMS
            for keys, values in center_dict.items():
                if student['center_code'] in values:
                    center_code = keys
                else:
                    print("not correct center code")     
            class_code = center_dict[center_code][2]
            student_name = student["student_name"]
            student_dob = student["dob"]
            parent_name = student["parent_name"]
            login_code = center_dict[center_code][1]
            print(class_code)
            print(student_name)
            print(center_code)
            ############################
        # Login Center
        windows = login.login_center(driver,login_code,center_code)
        admin_page = windows[0]
        center_page_window = windows[1]
        # Head to Student List
        driver.get('https://join.alab.edu.vn/member/memberList.do')
        sleep(2)
        # Searching ems_id
        # Choose ID Field
        WebDriverWait(driver=driver,timeout=10).until(
            EC.element_to_be_clickable(
                (By.XPATH,'//*[@id="eles-dash-main"]/div[3]/div[2]/div[2]/div[2]/select/option[2]'))
        ).click()
        sleep(2)
        # Fill teacher_id to search box
        driver.find_element("id","search").send_keys(ems_id)
        sleep(2)
        # Click on search button
        WebDriverWait(driver=driver,timeout=10).until(
            EC.element_to_be_clickable(
                (By.XPATH,'//*[@id="buttonSearch"]'))
        ).click()
        sleep(2)
        ### CHECKING: if the student id is not found, create student. Else if the student is found, then do not create student.
        search_results_rows = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr')
        search_results_cols = driver.find_elements(By.XPATH,'//*[@id="eles-dash-main"]/div[3]/table/tbody/tr/td')
        if((len(search_results_rows) == 1) & (len(search_results_cols)==1)):
            print("Student is not found")
            # Click on Batch Registration
            WebDriverWait(driver=driver,timeout=10).until(
                EC.element_to_be_clickable(
                    (By.XPATH,'//*[@id="batchRegBtn"]'))
            ).click()
            # Switching to popup batch registration window
            for handle in driver.window_handles:
                if ((handle != admin_page) & (handle != center_page_window)):
                    popup_batch_reg = handle
                    driver.switch_to.window(popup_batch_reg)
                    # NOW WE ARE IN POP-UP WINDOW
                    # Fill Batch Template and return its file name
                    file_name_template = self.fill_template(ems_id,student_name,student_dob,class_code,parent_name)
                    # Upload file
                    upload_input = driver.find_element(By.XPATH,'//*[@id="excelFile"]')
                    upload_input.send_keys(os.getcwd()+'\\'+ file_name_template)
                    sleep(2)
                    # Click on Register button
                    WebDriverWait(driver=driver,timeout=10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,'//*[@id="batchRegistBtn"]'))
                    ).click()
                    # Accept 1st alert message
                    driver.switch_to.alert.accept()
                    sleep(3)
                    # Checking 2nd alert message
                    alert_message = driver.switch_to.alert.text
                    message_num = message_dict[alert_message][0]
                    message = message_dict[alert_message][1]
                    match message_num:
                        case 1:
                            print('Student ID existed!!!!')
                            #Click Cancel
                            driver.switch_to.alert.dismiss()
                        case 2: 
                            print('Student ID empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 3: 
                            print('Student Password empty!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 4: 
                            print('Student Name empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 5: 
                            print('Class Code empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 6: 
                            print('Wrong Class Code empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 7: 
                            print('Parent Name empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 8: 
                            print('Parent Code empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 9: 
                            print('Parent Code with wrong format!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 10: 
                            print('Parent Password empty!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case 200: 
                            print('Create Student Succesfully!!!!!')
                            #Cick Ok
                            driver.switch_to.alert.accept()
                        case _:
                            print('Special Message Found!!!!!!')
                    # Delete file
                    os.remove(os.getcwd()+'\\'+file_name_template)
                    return {
                        "LMS_ID" : ems_id,
                        "Center" : center_code,
                        "Student_Name": student_name,
                        "Message_Num": message_num,
                        "Message": message
                    }
                else:
                    print("Not the right window")    
        elif((len(search_results_rows) == 1) & (len(search_results_cols)>1)):
            print("Student is found with 1 result")
            return {
                "LMS_ID": ems_id,
                "Center": center_code,
                "Student_Name" : student_name,
                "Message": "1 Student ID has been found!!!"
            }
        elif ((len(search_results_rows) > 1)):
            return {
                "LMS_ID": ems_id,
                "Center": center_code,
                "Student_Name" : student_name,
                "Message": "More than 1 Student ID has been found!!!"
            }
        else:
            print("Special Case!!!!!")

    def fill_template(self,student_id,student_name,student_dob,class_code,parent_name):
        now_string = date.datetime.now()
        date_time_string = now_string.strftime("%m-%d-%Y_%H-%M-%S")
        file_name = 'member-template-filled-{}.xls'.format(date_time_string)
        batch_template = pd.read_excel("member_template.xlsx", sheet_name="Sheet1")
        # Fill Student ID, DOB, Class Code, and Parent Name into template file
        batch_template['ID'] = student_id
        batch_template['Name'] = student_name
        batch_template['BirthDate'] = student_dob
        batch_template['Class code'] = class_code
        batch_template['parent name'] = parent_name
        # Save to new file
        batch_template.to_excel(file_name,index=False,engine='xlwt')
        return file_name
        
        
        
    

        
        