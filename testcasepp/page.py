# import sys
import time

# import openpyxl
from selenium.webdriver.support import expected_conditions as EC



# from autoit import autoit


import pyautogui
# import pywinauto
import os
# from openpyxl import load_workbook
# from pywin.framework import window
from pywinauto import application, Application
# from selenium.common.exceptions import StaleElementReferenceException, NoAlertPresentException, TimeoutException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from excelutils import writedata, foldercheck
from locators import BiPageLocators




class Home(object):

    def __init__(self, driver):
        self.driver = driver

    def selectfolder(self):
        name = self.driver.find_element(*BiPageLocators.folder)
        name.click()



    def selectsfolder(self):
        time.sleep(7)
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "dsFileManager"))
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "dlgIframe"))
        print("Iframe entered")

        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Switch to DOCUMENTS.png")
        time.sleep(7)

    def selects1folder(self):
        s1 = WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, ('//span[contains(.,"Project documents")]'))))
        time.sleep(7)
        if s1.is_displayed():
                print("Element found")
                s1.click()
        else:
                print("Element not found")
                self.driver.refresh()




        '''try:
            s1 = self.driver.find_element(*BiPageLocators.s1folder)
            s1.click()
        except:
            self.driver.refresh()
            '''



        time.sleep(5)
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Navigate to Proj Doc .png")

    def selects2folder(self, foldername):
        s2 = self.driver.find_element_by_xpath(BiPageLocators.s2folder.format(foldername))
        action = ActionChains(self.driver)
        action.context_click(s2).perform()
        time.sleep(2)
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\upload/replace.png")

    def selectdropdown(self, documentname):
        s3 = self.driver.find_element(*BiPageLocators.dropdown)
        s3.click()
        time.sleep(2)
        App = application.Application()  # Instantiate Application
        # The class used here without the window title, mainly to ensure compatibility
        App.connect(class_name='#32770', title='Open', visible_only=True,
                    found_index=0)  # Find pop-ups based on class_name
        App["Dialog"]["Edit1"].type_keys(documentname)
        print("fname printed")
        time.sleep(3)
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\File Upload.png")# Enter a value in the input box
        App["Dialog"]["Button1"].click()
        try:
            App["Dialog"]["Button1"].click()
        except:
            pass

        # autoit.control_send("[CLASS:#32770]", "Edit1", "C:\\Users\\hp\\Desktop\\image 7.jpg")
        # autoit.control_click("[Class:#32770]", "Button1")

    def selectcbox(self):
        time.sleep(5)
        s4 = self.driver.find_element(*BiPageLocators.cbox)
        self.driver.execute_script("arguments[0].click();", s4)
        time.sleep(5)
        print("Completed")
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Checkbox selected.png")
        time.sleep(7)

    def selects5(self):
        s6 = self.driver.find_element(*BiPageLocators.s5)
        print("xpath works")
        # self.driver.execute_script("arguments[0].click();", s6)
        time.sleep(7)
        s6.click()
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Click OK.png")
        print("Process completed")
        time.sleep(20)

    def searchpane(self,documentname):
        frame = documentname.split("\\")[-1]
        s0 = self.driver.find_element(*BiPageLocators.sp)
        s0.click()
        time.sleep(3)
        print("found searchpane")
        k1 = self.driver.find_element(*BiPageLocators.keyword)
        k1.send_keys(frame)
        #self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\File Keyword search.png")
        G2 = self.driver.find_element(*BiPageLocators.search)
        G2.click()

        time.sleep(5)
        filevalidate = self.driver.find_elements(*BiPageLocators.s9)
        print("")
        # self.driver.refresh()
        data = []
        print(data)
        for item in filevalidate:
            data.append(item.text)

        return data


class Importfold(object):

    def __init__(self, driver):
        if driver is None:
            driver = {}
        else:
            self.driver = driver


    def selectfolder(self):
        name = self.driver.find_element(*BiPageLocators.folder)
        name.click()


    def selectsfolder(self):
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "dsFileManager"))
        self.driver.switch_to.frame(self.driver.find_element(By.ID, "dlgIframe"))

        # self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Switch to DOCUMENTS.png")
        time.sleep(7)

    def selects1folder(self):


        s1 = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, ('//span[contains(.,"Project documents")]'))))
        time.sleep(7)
        if s1.is_displayed():
            print("Element found")

        else:
            print("Element not found")
            self.driver.refresh()

        action = ActionChains(self.driver)
        action.context_click(s1).perform()

    def  selectnew(self,newfoldname):
         n1=self.driver.find_element(*BiPageLocators.newfol)
         n1.click()
         n2=self.driver.find_element(*BiPageLocators.folname)
         n2.send_keys(newfoldname)
         n3=self.driver.find_element(*BiPageLocators.outclk)
         n3.click()

    def selectn4folder(self, newfoldname):
        n4 = self.driver.find_element_by_xpath(BiPageLocators.n4folder.format(newfoldname))
        pyautogui.click(x=365, y=545)
        #n4.click()
        time.sleep(2)

    def selectn6folder(self, newfoldname):
        n9 = self.driver.find_element_by_xpath(BiPageLocators.n4folder.format(newfoldname))

        action = ActionChains(self.driver)
        action.context_click(n9).perform()
        try:
            action.context_click(n9).perform()
        except:
            pass

        n5=self.driver.find_element(*BiPageLocators.impfile)
        n5.click()




    def importfolderstr(self, emptyfolderpath):
        main_page=self.driver.switch_to.window(self.driver.window_handles[1])
        i1=self.driver.find_element(*BiPageLocators.folstr)
        i1.click()
        time.sleep(3)
        path = emptyfolderpath
        isdir = os.path.isdir(path)



        # Workbook = openpyxl.load_workbook('..\\TEST data\\TestSteps Suite.xlsx')
        #
        # sheet = Workbook('ImportFolder')
        # sheet['D8'] = 'Folder is not available in respective path;Please recheck'
        # Workbook.save('..\\TEST data\\TestSteps Suite.xlsx')

        if isdir == True:
        # os.system(emptyfolderpath)
            pyautogui.write(emptyfolderpath)
        # pyautogui.press('Select Folder')
            pyautogui.press('enter',presses=2)
        else:
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', 8, 'Fail ')
            foldercheck('..\\TestSteps Suite.xlsx', 'ImportFolder', 8, 'Folder not available;Please recheck')


        time.sleep(2)
        print("Folder uploaded")

        pyautogui.click(x=651,y=240)

        i3=self.driver.find_element(*BiPageLocators.progtextbar)

        while i3.get_attribute('innerHTML') != '100%':
            print(i3.get_attribute('innerHTML'))
            time.sleep(5)

        time.sleep(10)

        i2 = self.driver.find_element(*BiPageLocators.close)
        self.driver.execute_script("arguments[0].click();", i2)








