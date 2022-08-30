from selenium.webdriver.common.by import By


# from selenium.webdriver.common.keys import Keysa

class BiPageLocators(object):
    # folder = (By.XPATH, ("//*[@id='ucMenuModule_menu3']/a"))
    folder = (By.XPATH,('//a[contains(text(),"Documents")]'))
    #folder = (By.LINK_TEXT, ("DOCUMENTS"))
    s1folder = (By.XPATH, ('//span[contains(.,"Project documents")]'))
    s2folder = '//div[@role="presentation"]//span[text()= "{}"]'
    #s2folder = (By.XPATH,('// span[contains(., "{}")]'))
    dropdown = (By.XPATH, ('//div[@role = "menubar"]//a[5]'))
    cbox = (By.XPATH, ('//input[@id="chk0"]'))
    s5 = (By.XPATH, ('//div/div/form/footer/div/div/div/div/button[2]'))
    sp = (By.XPATH,('//button[contains(.,"Show search pane")]'))
    keyword=(By.XPATH,('//div[@view_id="dsSearchToolBar"]//input[@type="text"]'))
    search =(By.XPATH,('//button[contains(.,"Search")]'))
    s9 = (By.XPATH, ('//div[@view_id="modeViews"]//div[@role="gridcell"][@aria-colindex="2"]'))
    newfol=(By.XPATH,('//a[contains(text(),"New folder")]'))
    folname=(By.XPATH,('(//input[@type="text"])[5]'))
    outclk=(By.XPATH,('//div[2]/div/div/div[2]'))
    n4folder='//div[2]/div/div/span'
    impfile=(By.XPATH,('//a[contains(text(),"Import folders and files to an empty folder")]'))
    folstr=(By.XPATH,('//*[@id="form1"]//input[@id="buttonSelect"][@value="Select folder structure"]'))
    progtextbar=(By.XPATH,('//div[@id="progressTotalBar"]'))
    close=(By.XPATH,('//input[@id="buttonDone"][@value="Close"]'))

#