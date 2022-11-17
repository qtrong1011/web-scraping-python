from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from lmsJobs import DriverBuilder, LoginAdminLMS, ResetPassword





##### MAIN DRIVER #########
if __name__ == '__main__':
   AO_driver = DriverBuilder().get_driver()
   teacher_driver = DriverBuilder().get_driver()
   # LoginAdminLMS().login_mag_admin(driver)
   password_AO = "Alab2022"
   password_teacher = "alab202210"
   reset_pwd_AO = ResetPassword()
   reset_pwd_teacher = ResetPassword()
   print('------Starting Reset Password for AOs-----')
   reset_pwd_AO.reset_pwd_AO(AO_driver,password_AO)
   print('------Finshed Reset Password for AOs-----')
   print('------Starting Reset Password for Teachers-----')
   reset_pwd_teacher.reset_pwd_teachers(teacher_driver,password_teacher)
   print('------Finshed Reset Password for Teachers-------')
