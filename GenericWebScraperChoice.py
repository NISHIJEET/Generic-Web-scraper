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
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
import pandas as pd
import xlrd
import xlwings as xw

#// Defines the preference so that there is no prompt
#// for any download
prefs_downloadPrompt = { "download.prompt_for_download" : False}

list_DataFrameCurrentRecord=[]
dict_dataFrames={}

#// Defines the options that need to be passed for our chrome object
def setChromeOptions(prefs_downloadPrompt):
    chromeOptions = webdriver.ChromeOptions()

    #// set the required preferences
    chromeOptions.add_experimental_option("prefs",prefs_downloadPrompt)
    return chromeOptions


#// creating the driver that will help our code to interact with the chrome object
def createDriver(chromeOptions,ChromeDriver_Path):

    #// Setting the chrome settings defined earlier
    caps = DesiredCapabilities().CHROME

    #// page load strategy is basically until what time should the could wait for the page to load
    #// Right now, it starts to execute as soon as the page becomes responsiev and 
    #// dosent wait for the page to completely load
    caps["pageLoadStrategy"] = "none"
    driver = webdriver.Chrome(desired_capabilities=caps,chrome_options=chromeOptions,executable_path=ChromeDriver_Path)
    return driver


#// Gets the input table, the algorithm table and the control sheet 
def GetAllRequiredDetails():

    #// getting the list of property and checkin checkoput detail table
    dict_dataFrames["InputTable" + str(len(dict_dataFrames)+1)] = pd.read_excel(excel_file,sheet_name="Tables")
    #// getting the control table that contains the scraping algorithm
    dict_dataFrames["ControlTable"] = pd.read_excel(excel_file,sheet_name="Control Table")
    #// getting the details of each action
    dict_dataFrames["MetaData"] = pd.read_excel(excel_file,sheet_name="Control")


#// Main  method that controls the flow of execution
def ControllerMain(s_driverPath,s_outputFilename,s_workBookName):

    #// get the toll that controls the whole process
    global excel_file

    excel_file = s_workBookName
    
    #// set the chrome options that need to be passed into the driver
    chromeOptions =  setChromeOptions(prefs_downloadPrompt)
    #// Create the driver that runs the Chrome
    #// chrome window will pop up after this step
    driver = createDriver(chromeOptions,s_driverPath)

    #// getting all the required details to run the process
    GetAllRequiredDetails()

    global rowVal
    global list_returnedValues

    #// Main array that will store the list of all returned values
    list_returnedValues = []
    int_records = 0

    #// looping on all the properties and their corresponding 
    #// checkin date and checkout date
    for indexVal, rowVal in df_inputTable.iterrows(): 

        #// try block here catches any exception/error at property level and
        #// generate an error statement corresponding to that property
        try:
            #// for each property we will loop on the defined set of procedures
            #// that are defined in the algorithm table
            for index, row in df_control.iterrows():

                #// s_action stores the ActionToBePerformed eg. Navigate, click etc
                s_action = row['Action']

                #// filter in meta data for the action in this step
                df_tempMetaData = df_metaData[df_metaData['Actions'] == s_action]
                #// Drop all records with NaN
                df_tempMetaData.dropna
                #// In dataFrame, each data row has an index with which we identify it
                #// we need to reset the index after filtering in data
                #// inplace=true measn that it wont retur nany data
                df_tempMetaData.reset_index(inplace = True)

                #// stores if we require to save the return value or not
                #// if we do, then it contains  TRUE||s_tag 
                s_saveField = str(df_tempMetaData.ix[0,'Save Field'])
                #// what field contaisn the details of action in the main controller table
                s_ActionField = df_tempMetaData.ix[0,'Action Field']
                print(s_ActionField)
                #// the value with whoch the action needs to be performed
                s_ActionValue = df_control.ix[index,s_ActionField]
                print(s_ActionValue)
                #// the python function that needs to be called for that action
                s_functionName = df_tempMetaData.ix[0,'Python Function']
                #// call the python function for that action and pass in the details 
                #// requiored. Architecture made in such a way that it always requires
                #// only two arguements. They retuen a value that can be stored
                s_saveValue = globals()[s_functionName](driver,s_ActionValue)
                print("OUTPUT>>>" + s_saveValue)

                #// saving the returned value in returnlsit and print the list to 
                #// output file
                if str("TRUE") in s_saveField and "<>" not in s_saveValue:
                    s_actionTag = s_saveField.split("::")[1]
                    list_returnedValues.append(s_actionTag + "||" + s_saveValue)
                    print("Save::Added to list")
                    if len(list_returnedValues) > int_records:
                        df_temp = pd.DataFrame(list_returnedValues)
                        int_records = int_records + 1
                        df_temp.to_excel(s_outputFilename, header=False, index=False)
                        print("Output made to excel")
        except:
            #// error return value that needs to be generated for a property
            list_returnedValues.append("Error||" + rowVal['Property Code'] + "||" + "Automation error! Please check for any change in website or logic.")
            print("Error::Added to list")
            df_temp = pd.DataFrame(list_returnedValues)
            int_records = int_records + 1
            df_temp.to_excel(s_outputFilename, header=False, index=False)
            print("Output made to excel")

        print("----------------------------")
        print(list_returnedValues)


#// Picks in the checkin date
def Action_DatePickerCheckOut(driver,s_funcnName):
    try:
        date_picker_choice(driver,CheckOutDay,CheckOutMonthYear)
        return "date picked"
    except StaleElementReferenceException:
        pass
        return "date error"

#// Picks in the checkout date
def Action_DatePickerCheckIn(driver,s_funcnName):
    try:
        date_picker_choice(driver,CheckInDay,CheckInMonthYear)
        return "date picked"
    except StaleElementReferenceException:
        pass
        return "date error"

#// Navigates to a URL
def Action_Navigate(driver,URL):
    try:
        URL = ReplaceWithTableVal(URL)
        driver.get(URL)
        return "Url navigated"
    except:
        pass
        return "URL error"

#// Click on an elemetn
def Action_Click(driver,XPATH):
    try:    
        driver.find_element(By.XPATH,XPATH).click()
        return "Clicked"
    except:
        pass
        return "Element not found"

#// Replaces the <<asd::asd>> format with value from the table
def ReplaceWithTableVal(s_textToReplace):
    if "<<" in s_textToReplace:
        s_textToReplace = s_textToReplace.replace("<<","")
        s_textToReplace = s_textToReplace.replace(">>","")
        s_textToReplace = rowVal[s_textToReplace.split("::")[1]]

    return s_textToReplace

#// send keys to a particular element
#// receives XPATH as XPATH||Keyvalue
def Action_SendKeys(driver,XPATH):
    s_keyToType = XPATH.split("||")[1]

    s_keyToType = ReplaceWithTableVal(s_keyToType)    

    XPATH = XPATH.split("||")[0]
    try:
        driver.find_element(By.XPATH,XPATH).send_keys(s_keyToType)
        return "Entered"
    except:
        pass
        return "Element not found"

#// sleep for given time
def Action_Sleep(driver,time):
    sleep(time)
    return "Slept"

#// Pickup value from element
def Action_PickUp(driver,XPATH):
    s_tag = XPATH.split("||")[1]
    s_tag = ReplaceWithTableVal(s_tag) 

    XPATH = XPATH.split("||")[0]
    try:
        return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
    except:
        return "Element not found"

#// Clear value from element
def Action_Clear(driver,XPATH):
    try:
        driver.find_element(By.XPATH,XPATH).clear()
        return "cleared"
    except:
        return "Element not found"

#// checks if an elemet exists ot not
def Action_CheckIfExists(driver,XPATH):
    try:
            driver.find_element_by_xpath(XPATH)
    except NoSuchElementException:
            return False

    return True
 
def Action_CustomIfPresent(driver,XPATH):
    s_tag = XPATH.split("||")[1]
    s_tag = ReplaceWithTableVal(s_tag) 
    s_custom = XPATH.split("||")[2]
    s_custom = ReplaceWithTableVal(s_custom) 
    XPATH = XPATH.split("||")[0]
    
    if Action_CheckIfExists(driver,XPATH):
        return s_tag + "||" + s_custom
    else:
        return "<>"


def Action_PickUpIfAvailable(driver,XPATH):
    s_tag = XPATH.split("||")[1]
    s_tag = ReplaceWithTableVal(s_tag) 
    XPATH = XPATH.split("||")[0]

    if Action_CheckIfExists(driver,XPATH):
        return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
    else:
        return "<>"

def Action_PickUpIfAvailableElse(driver,XPATH):
    s_tag = XPATH.split("||")[1]
    s_tag = ReplaceWithTableVal(s_tag) 
    XPATH = XPATH.split("||")[0]
    s_elseText = XPATH.split("||")[2]
    s_elseText = ReplaceWithTableVal(s_elseText) 

    if Action_CheckIfExists(driver,XPATH):
        return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
    else:
        return s_tag + "||" + s_elseText

def Action_PickIfNotNull(driver,XPATH):
    s_tag = XPATH.split("||")[1]
    s_tag = ReplaceWithTableVal(s_tag) 
    XPATH = XPATH.split("||")[0]

    if Action_CheckIfExists(driver,XPATH):
        if driver.find_element(By.XPATH,XPATH).text.strip() == "":
            return "<>"
        else:
            return s_tag + "||" + driver.find_element(By.XPATH,XPATH).text
    else:
        return "<>"

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
        time.sleep(1)
        date_picker_choice(driver,Day,MonthYear)


#// get the workbook details and open an instance
wb = xw.Book("GenericWebCrawler.xlsm")
val_chromepath = wb.sheets['KeyValues'].range("KeyValues_ChromedriverPath").value
val_outputDir = wb.sheets['KeyValues'].range("KeyValues_OutputFilePath").value
val_mainPath = wb.sheets['KeyValues'].range("KeyValues_MainPath").value

ControllerMain(val_chromepath,val_outputDir,val_mainPath)