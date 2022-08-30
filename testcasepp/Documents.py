
import time
import unittest
from selenium import webdriver
from selenium.webdriver.support.select import Select
from webdriver_manager.microsoft import EdgeChromiumDriverManager

from page import Home, Importfold
from msedge.selenium_tools import Edge, EdgeOptions
from excelutils import getlogindata, gettabname,gettestdatadoc,gettestdataimpfol, writedata, removedata,writeerror,projectno








class Search(unittest.TestCase):

    def setUp(self):
        removedata('..\\TestSteps Suite.xlsx', 'Sheet1')
        removedata('..\\TestSteps Suite.xlsx', 'ImportFolder')
        # self.driver = webdriver.Edge(executable_path=DRIVERPATH)
        # self.driver = webdriver.Chrome(executable_path=DRIVERPATH)
        # self.driver = webdriver.Chrome(executable_path="C:\\Users\\hp\\edge\\msedgedriver.exe", options=c)


        #self.driver = Edge(executable_path=driver_path, options=options)
        #self.driver=getdriverconfig(Edge(executable_path=(r'..\\Config\\configloc.xlsx','config'), options=options))
        #self.driver.get("https://wa-test-cmn-projectportal.ase02.ncc.info//")

        driver_path = EdgeChromiumDriverManager().install()
        options = EdgeOptions()
        options.use_chromium = True
        options.add_experimental_option("detach", True)
        options.add_argument("-inprivate")
        self.driver = Edge(executable_path=driver_path, options=options)
        self.driver.get("https://wa-test-cmn-projectportal.ase02.ncc.info//")
        #self.driver.get("https://wa-test-cmn-projectportal.ase02.ncc.info//")
        self.driver.implicitly_wait(10)
        self.driver.maximize_window()
        time.sleep(5)

        time.sleep(30)

        #self.driver.find_element_by_id("idSIButton9").click()
        #time.sleep(5)
        #self.driver.find_element_by_name("passwd").send_keys(password)
        #time.sleep(5)
        #self.driver.find_element_by_id("idSIButton9").click()

        time.sleep(7)
        self.driver.save_screenshot("C:\\Users\\hp\\Desktop\\Test Execution Snapshots\\Login Snapshot.png")
        writedata('..\\TestSteps Suite.xlsx', 'Sheet1', 2, 'PASS')
        writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', 2, 'PASS')

        self.driver.find_element_by_id("idSIButton9").click()
        time.sleep(3)
        self.driver.find_element_by_xpath("//a[contains(text(),'select')]").click()

        selectprojecno = projectno('..\\TestData.xlsx', 'TestCase1')

        self.driver.find_element_by_xpath(f"//span[contains(.,'{selectprojecno}')]").click()
        time.sleep(3)
        #self.driver.find_element_by_xpath("//span[contains(.,'"+selectprojecno+"')]")
        #select.select_by_value(str(selectprojecno))
        print(selectprojecno)

        writedata('..\\TestSteps Suite.xlsx', 'Sheet1', 3, 'PASS')
        writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', 3, 'PASS')


    def test_login(self):
        tabname = gettabname('..\\TestData.xlsx', 'TestCase1')
        if tabname=="Documents":
            self.document_tab()
        else:
            self.import_folder()

    def document_tab(self):

        stepcount = 4
        try:
            documentname,foldername= gettestdatadoc('..\\TestData.xlsx', 'TestCase1')
            home = Home(self.driver)
            home.selectfolder()
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            time.sleep(3)
            home.selectsfolder()
            # time.sleep(3)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            home.selects1folder()
            time.sleep(3)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            home.selects2folder(foldername)
            time.sleep(3)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            home.selectdropdown(documentname)
            time.sleep(5)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            home.selectcbox()
            time.sleep(5)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            stepcount += 1
            home.selects5()
            time.sleep(5)
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'PASS')
            home.searchpane(documentname)
            time.sleep(3)

        except Exception as r:
            writedata('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount, 'FAIL')
            writeerror('..\\TestSteps Suite.xlsx', 'Sheet1', stepcount,  str(r))
            self.fail()


    def import_folder(self):
        stepcount = 4
        try:
            newfoldname,emptyfolderpath = gettestdataimpfol('..\\TestData.xlsx', 'TestCase1')
            impfol= Importfold(self.driver)
            impfol.selectfolder()
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'PASS')
            stepcount += 1
            time.sleep(2)
            impfol.selectsfolder()
            time.sleep(2)
            impfol.selects1folder()
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'PASS')
            stepcount += 1
            time.sleep(2)
            impfol.selectnew(newfoldname)
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'PASS')
            stepcount += 1
            time.sleep(2)
            impfol.selectn4folder(newfoldname)

            time.sleep(2)
            impfol.selectn6folder(newfoldname)
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'PASS')
            stepcount += 1
            time.sleep(2)
            impfol.importfolderstr(emptyfolderpath)
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'PASS')
            time.sleep(2)

        except Exception as e:
            writedata('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, 'FAIL')
            writeerror('..\\TestSteps Suite.xlsx', 'ImportFolder', stepcount, str(e))
            self.fail()



    def tearDown(self):
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()



