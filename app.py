from selenium import webdriver # (1) login 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import numpy as np
import openpyxl

driver = webdriver.Chrome()
driver.get("https://az3.ondemand.esker.com/ondemand/webaccess/asf/home.aspx")
driver.maximize_window()
time.sleep(1)

driver.find_element(By.XPATH, '//*[@id="ctl03_tbUser"]').send_keys("john.tan@sh-cogent.com.sg")
driver.find_element(By.XPATH, '//*[@id="ctl03_tbPassword"]').send_keys("Esker3838")
driver.find_element(By.XPATH, '//*[@id="ctl03_btnSubmitLogin"]').click()
time.sleep(0.5)

def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

x_path_hover = '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div'
hover(driver, x_path_hover)
time.sleep(1)

try:
    drop_down=driver.find_element(By.XPATH, '//*[@id="DOCUMENT_TAB_100872215"]/a/div[2]').click()
    time.sleep(1)
except Exception as e:
    print(e) #VENDOR INVOICES
#driver.find_element(By.XPATH, '//*[@id="mainMenuBar"]/td/table/tbody/tr/td[36]/a/div') # 1(a)

path_vendor = "https://raw.githubusercontent.com/JohnTan38/Best-README/master/"
vendor_data = pd.read_excel(path_vendor+ 'Vendor_Data.xlsx', sheet_name='Sheet1', engine='openpyxl')
vendor_payment_reference = pd.read_excel(path_vendor+ 'Vendor_Payment_Reference.xlsx', sheet_name='vendor_payment', engine='openpyxl')

vendor_data = vendor_data[['Vendor Number', 'Name1']]
vendor_payment_reference = vendor_payment_reference[['Vendor', 'Payment reference', 'Clrng doc.', 'Pstng Date']]
vendor_data.rename(columns={'Vendor Number': 'Vendor_Number'}, inplace=True)
vendor_payment_reference.rename(columns={'Payment reference': 'Payment_reference', 'Clrng doc.': 'Clrng_doc', 'Pstng Date': 'Pstng_Date'}, inplace=True)
vendor_payment_reference = vendor_payment_reference.dropna(subset=['Vendor'])

def filter_vendor_number(df):
    df['Vendor_Number'] = df['Vendor_Number'].astype(str) # Ensure 'Vendor_Number' is of string type
    # Remove rows where 'Vendor_Number' is blank, None, or NaN
    df = df[df['Vendor_Number'].astype(str).str.strip().notna() & df['Vendor_Number'].astype(str).ne('nan')]
    df = df[df['Vendor_Number'].str.startswith('1000')] # Select rows where 'Vendor_Number' starts with '1000'
    return df

def filter_payment_reference(df):
    df['Clrng_doc'] = df['Clrng_doc'].astype(str)
    df = df[df['Payment_reference'].astype(str).str.strip().notna() & df['Payment_reference'].astype(str).ne('nan')]
    df = df[df['Clrng_doc'].astype(str).str.strip().notna() & df['Clrng_doc'].astype(str).ne('nan')]
    df = df[df['Clrng_doc'].str.startswith('15')]
    return df

from datetime import datetime, timedelta
def filter_dates(df):
    df['Pstng_Date'] = pd.to_datetime(df['Pstng_Date']) # Convert 'Doc_Date' to datetime format    
    sixty_days_ago = datetime.now() - timedelta(days=60) # Get the date 60 days ago from today        
    df = df[df['Pstng_Date'] > sixty_days_ago] # Select rows where 'Doc date' is within the past 60 days from today
    return df

def filter_top_vendor(df, column_name, list_top_vendor):
    return df[df[column_name].isin(list_top_vendor)]

top_vendors = [1000366487,1000286720,1000169276,1000152171] # 1000366487 1000286720 1000169276 1000152171
vendor_data = filter_vendor_number(vendor_data)
vendor_payment_reference = filter_payment_reference(vendor_payment_reference)
vendor_payment_reference_sixtydays = filter_dates(filter_payment_reference(vendor_payment_reference))

dict_vendor_data = vendor_data.set_index('Vendor_Number').to_dict()['Name1'] #%timeit

def get_vendor_name(top_vendor):
     return dict_vendor_data.get(str(top_vendor))

def extract_invoice_numbers(driver, x_path_invoice):
    list_invoice_numbers = [] # Initialize an empty list to store the invoice numbers
    # Get the number of rows in the table
    rows = driver.find_elements(By.XPATH, '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr')
    # Iterate over each row
    for i in range(2, len(rows)-1):
        # Construct the XPath for the current row
        x_path_invoice_current = x_path_invoice.replace('tr[2]', f'tr[{i}]')

        # Find the element and get its innerHTML attribute
        try:
            invoice_number = driver.find_element(By.XPATH, x_path_invoice_current).get_attribute("innerHTML")
        except Exception as e:
            print(e)
        list_invoice_numbers.append(invoice_number)
    return list_invoice_numbers #get list invoice numbers on current page

def update_invoices_index(list_invoice, dict_update_vendor_payment_reference):
    result = []
    for key, value in dict_update_vendor_payment_reference.items():
        if str(key) in list_invoice:
            invoice_number_to_update = str(key)
            index_invoice_to_update = list_invoice.index(str(key))
            result.append((invoice_number_to_update, value, index_invoice_to_update))
    return result


for top_vendor in top_vendors:
    try:
        company_code = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl02_ddl1_tagify"]/span')
        company_code.send_keys("SG77") #(2) input company code SG77
        time.sleep(0.3)
        actions = ActionChains(driver)
        num_tab = 6
        actions.send_keys(Keys.TAB*num_tab).perform()
    except Exception as e:
        print(e)
    update_vendor_payment_reference = filter_top_vendor(vendor_payment_reference_sixtydays, 'Vendor', [top_vendor])
    dict_update_vendor_payment_reference = update_vendor_payment_reference.set_index('Payment_reference').to_dict()['Clrng_doc']
    vendor_name = get_vendor_name(top_vendor)    
    try:
        vendor_name_0 = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl03_ddl1_tagify"]/span')
        vendor_name_0.send_keys(vendor_name)
        time.sleep(0.5)
        actions.send_keys(Keys.ENTER).perform()
    except Exception as e:
        print(e)

    #tuple_update_payment_reference = update_invoices_index(list_invoice_numbers_on_page, dict_update_vendor_payment_reference)
    for key,value in dict_update_vendor_payment_reference.items():
        invoice_number_to_update = key
        payment_reference_to_update = value
        x_path_invoice = '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr[2]/td[7]/a'
        list_invoice_numbers_on_page = extract_invoice_numbers(driver, x_path_invoice) #get list invoice numbers on current page
        tuple_update_payment_reference = update_invoices_index(list_invoice_numbers_on_page, dict_update_vendor_payment_reference)
        tuple_update_payment_reference = [('202402556', '1500002236.0', 5), ('202402558', '1500002236.0', 4),
                                          ('CR408557', '1500002338.0', 1), ('CR407254', '1500002238.0', 0), 
                                          ('72351', '1500002392.0', 2), ('2176615', '1500002310.0', 4)]#, [('202402556', '5100003086.0', 5), ('202402558', '5100003089.0', 4), ('72351', '5100002392.0', 2)] ##
 
        ## get index
        for tuple_elem in tuple_update_payment_reference:
            if str(key) == tuple_elem[0]:
                index_invoice_toClick = tuple_elem[2]
                payment_reference_to_update = tuple_elem[1].replace('.0', '')
                
                x_path_invoice_toClick = '//*[@id="tpl_ih_adminList_MainGrid"]/tbody/tr[' + str(index_invoice_toClick+2) + ']/td[7]/a'
                invoice_toClick = tuple_elem[0]
                print(invoice_toClick, x_path_invoice_toClick)
                        
                try:
                    #driver.find_element(By.XPATH, x_path_invoice_toClick).click()
                    WebDriverWait(driver,3).until(EC.element_to_be_clickable((By.LINK_TEXT, invoice_toClick))).click()
                except Exception as e:
                    print(e)
                time.sleep(5)
                try:
                    enter_payment_details=driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[5]/span').click() #Enter payment details btn
                except Exception as e:
                    print(e)
                time.sleep(3)

                try:
                    driver.find_element(By.XPATH, '//*[@id="Payment_pane_eskCtrlBorder_content"]/div/div/table/tbody/tr[3]/td[2]/div/input').click()
                    driver.find_element(By.XPATH, '//*[@id="Payment_pane_eskCtrlBorder_content"]/div/div/table/tbody/tr[3]/td[2]/div/input').send_keys(payment_reference_to_update)
                    time.sleep(5) #input payment reference
                except Exception as e:
                    print(e)

                try:
                    save_click = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[1]/span')#.click() # Save btn
                    #WebDriverWait(driver,3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form-footer"]/div[1]/a[1]/span'))).click()
                except Exception as e:
                    print(e)
                #back to page 1
                try:
                    quit_click_1 = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[2]/span').click()
                except Exception as e:
                    print(e)
                time.sleep(5)
                try:
                    quit_click_2 = driver.find_element(By.XPATH, '//*[@id="form-footer"]/div[1]/a[8]/span').click()
                except Exception as e:
                    print(e)
                time.sleep(5)
                try:
                    company_code = driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl02_ddl1_tagify"]/span')
                    company_code.send_keys("SG77") #(2) input Company code SG73, SG77, SG81, SG82, SG83, SG88
                    time.sleep(1)
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.TAB*6).perform()
                    driver.find_element(By.XPATH, '//*[@id="tpl_ih_adminList_displayedFilters_ctl00_ctl03_ddl1_tagify"]/span').send_keys(vendor_name)
                    actions.send_keys(Keys.ENTER).perform()
                    time.sleep(1)
                except Exception as e:
                    print(e)
            else:
                None

