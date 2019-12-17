from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException,NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pandas as pd
import xlrd

ChromeDriver_Path = 'C:\\Previous Downloads\\Setup\\chromedriver_win32\\chromedriver.exe'
prefs_downloadPrompt = { "download.prompt_for_download" : False }

list_DataFrameCurrentRecord=[]
dict_dataFrames={}
#// keep on checking the current while loop criteria
int_FlagWhileLoop=0
str_WhileLoopCriteria=""
excel_file = "GenericWebCrawler_NJ_v0.1.xlsx"

#// Defines the options that need to be passed for our chrome object
def setChromeOptions(prefs_downloadPrompt):
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_experimental_option("prefs",prefs_downloadPrompt)
    return chromeOptions


#// creating the driver that will help our code to interact with the chrome object
def createDriver(chromeOptions,ChromeDriver_Path):

    #// Setting the chrome settings defined earlier
    driver = webdriver.Chrome(chrome_options=chromeOptions,executable_path=ChromeDriver_Path)
    return driver



def GetAllRequiredDetails():

    dict_dataFrames["InputTable" + str(len(dict_dataFrames)+1)] = pd.read_excel(excel_file,sheetname="Tables")
    dict_dataFrames["ControlTable"] = pd.read_excel(excel_file,sheetname="Control Table")
    dict_dataFrames["MetaData"] = pd.read_excel(excel_file,sheetname="Control")

def ExecuteControlTableCurrentRecord(s_Action,row_index):
    df_control = dict_dataFrames.get("ControlTable")



def ControllerMain():

    chromeOptions =  setChromeOptions(prefs_downloadPrompt)
    driver = createDriver(chromeOptions,ChromeDriver_Path)

    GetAllRequiredDetails()

    df_control = dict_dataFrames.get("ControlTable")
    df_metaData = dict_dataFrames.get("MetaData")
    print(len(dict_dataFrames))
    print(len(df_control.index))
    
    for index, row in df_control.iterrows():
        s_action = row['Action']
        df_tempMetaData = df_metaData[df_metaData['Actions'] == s_action]
        df_tempMetaData.dropna
        df_tempMetaData.reset_index(inplace = True)
        s_saveField = df_tempMetaData.ix[0,'Save Field']
        s_ActionField = df_tempMetaData.ix[0,'Action Field']
        print(s_ActionField)
        s_ActionValue = df_control.ix[index,s_ActionField]
        print(s_ActionValue)
        s_functionName = df_tempMetaData.ix[0,'Python Function']
        s_saveValue = globals()[s_functionName](driver,s_ActionValue)
        print("OUTPUT>>>" + s_saveValue)

        
def Action_Navigate(driver,URL):
        driver.get(URL)
        return "Url navigated"

def Action_Click(driver,XPATH):
        driver.find_element(By.XPATH,XPATH).click()
        return "Clicked"

def Action_SendKeys(driver,XPATH):
        s_keyToType = XPATH.split("||")[1]
        XPATH = XPATH.split("||")[0]
        driver.find_element(By.XPATH,XPATH).send_keys(s_keyToType)
        return "Entered"

def Action_Sleep(driver,time):
        sleep(time)
        return "Slept"

def Action_PickUp(driver,XPATH):
        s_tag = XPATH.split("||")[1]
        XPATH = XPATH.split("||")[0]
        return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text

def Action_Clear(driver,XPATH):
        return driver.find_element(By.XPATH,XPATH).clear()

def Action_CheckIfExists(driver,XPATH):
        try:
                driver.find_element_by_xpath(XPATH)
        except NoSuchElementException:
                return False

        return True

def Action_PickUpIfAvailable(driver,XPATH):
        s_tag = XPATH.split("||")[1]
        XPATH = XPATH.split("||")[0]

        if Action_CheckIfExists(driver,XPATH):
                return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
        else:
                return "<>"

def Action_PickUpIfAvailableElse(driver,XPATH):
        s_tag = XPATH.split("||")[1]
        XPATH = XPATH.split("||")[0]
        s_elseText = XPATH.split("||")[2]

        if Action_CheckIfExists(driver,XPATH):
                return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
        else:
                return s_tag + "||" + s_elseText


def date_picker_choice(driver,Day,MonthYear):
    month = driver.find_element(By.XPATH,'//table/thead/tr/th[@colspan="5"]')
    #print(month.text)
    if month.text == MonthYear:
        #print("--------------------------")
        days = driver.find_elements(By.XPATH,'//div[@class="uib-datepicker"]/table/tbody/tr/td/button/span')
        flag = 0
        for day in days:
            if (flag==0 and day.text == '1'):
                flag=1
            elif(flag==1 and day.text == '1'):
                flag = 0
        
            if (flag ==1 and day.text==Day):
                #print("Clicking day")
                day.click()
                return
    else:
        driver.find_element(By.XPATH,'//div[@class="uib-datepicker"]/table/thead/tr/th[3]').click()
        time.sleep(2)
        date_picker_choice(driver,Day,MonthYear)

ControllerMain()