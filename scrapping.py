import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import ExcelUtils
import pandas as pd

chrome_path = Service(ChromeDriverManager().install())
option = webdriver.ChromeOptions()
option.add_argument("--start-maximized")
option.add_argument("--disable-geolocation")
option.add_argument("--ignore-certificate-errors")
option.add_argument("--disable-popup-blocking")
option.add_argument("--disable-translate")
driver = webdriver.Chrome(service=chrome_path, options=option)

file = "InputData.xlsx"
driver.get("http://environmentclearance.nic.in/searchproposal.aspx")
driver.implicitly_wait(20)

driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_RadioButtonList1_1").click()
                       
# sheet_rows = ExcelUtils.get_row_count(file, "Watch List")
# sheet_columns = ExcelUtils.get_column_count(file, "Watch List")

list_of_company = []

# for r in range(2, sheet_rows + 1):
#     input_data = ExcelUtils.read_data(file, "Watch List", r, 2)
#     list_of_company.append(input_data)

tmp_word_list = ["SARDA ENERGY & MINERALS LTD",
"NTPC Limited",
"Yasho Industries Ltd",
"VARDHMAN SPECIAL STEELS LTD",
"Tatva Chintan Pharma Chem Ltd"]

for tmp_word in tmp_word_list:
# tmp_word = "BHANSALI ENGINEERING POLYMERS LIMITED"
    search = (ExcelUtils.search_text_combination(tmp_word.split()))
    search_input = search[::-1]

    for i in search_input:
        temp = i
        search_area = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_textbox2")
        search_area.clear()
        search_area.send_keys(temp)
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btn").click()
        driver.refresh()
        rows = len(driver.find_elements(By.XPATH, '//table[@class="table Grid1"]/tbody/tr'))
        # print(rows)
        cols = len(driver.find_elements(By.XPATH, '//table[@class="table Grid1"]/tbody/tr/th'))
        # print(cols)
        for row in range(2, rows + 1):
            for col in range(1, cols + 1):
                if driver.find_elements(By.XPATH,
                '//*[@id="ctl00_ContentPlaceHolder1_grdevents"]/tbody/tr[' + str(row) + ']/td')[0].text == "No Records Found":
                    break
                else:
                    result = driver.find_elements(By.XPATH,
                                            '//*[@id="ctl00_ContentPlaceHolder1_grdevents"]/tbody/tr[' + str(
                                                row) + ']/td[5]')[0].text
                final_tmp = tmp_word.lower().replace("and","&")
                final_tmp = final_tmp.lower().replace("limited","ltd")
                final_result = result.lower().replace("and","&")
                final_result = final_result.lower().replace("limited","ltd")
                if final_tmp in final_result.lower():
                    if rows > 2:
                        click_item = driver.find_elements(By.XPATH,
                                                    '//*[@id="ctl00_ContentPlaceHolder1_grdevents"]/tbody/tr[' + str(
                                                        row) + ']/td[7]/headerstyle/itemstyle/headerstyle/itemtemplate/a/img')
                    else:
                        click_item = driver.find_elements(By.XPATH,
                                                    '//*[@id="ctl00_ContentPlaceHolder1_grdevents"]/tbody/tr/td/headerstyle/itemstyle/headerstyle/itemtemplate/a/img')
                    click_item[0].click()
                    driver.implicitly_wait(10)
                    outer_rows = len(driver.find_elements(By.XPATH, "//*[@class='table Grid1']/tbody/tr"))
                    outer_cols = len(driver.find_elements(By.XPATH, "//*[@class='table Grid1']/tbody/tr/th")) - 4
                    header_list = []
                    table_data_list = []
                    for r in range(2, outer_rows + 1):
                        for c in range(1, outer_cols + 1):
                            header = \
                                driver.find_elements(By.XPATH,
                                                    "//*[@class='table Grid1']/tbody/tr[1]/th[" + str(c) + "]")[
                                    0].text
                            header_list.append(header)
                            table_data = driver.find_elements(By.XPATH,
                                                            '//*[@class="table Grid1"]/tbody/tr[' + str(
                                                                r) + ']/td[' + str(c) + ']')[0].text
                            table_data_list.append(table_data)
                        dictionary_1 = {header_list[i]: table_data_list[i] for i in range(len(header_list))}
                        # print(dictionary_1)
                        imageLinks = driver.find_elements(By.XPATH,
                                                    '//*[@class="table Grid1"]/tbody/tr/td[8]//a')
                        # imageNames = []
                        # for element in imageLinks:
                        #     imageNames.append(element.get_attribute("src"))
                        # print(imageNames)
                        file_key = []
                        file_link = []
                        for link in imageLinks:
                            f = link.get_attribute("title")
                            file_key.append(f)
                            lst = link.get_attribute("href")
                            file_link.append(lst)
                            for idx, item in enumerate(file_key):
                                if "TOR" in item:
                                    file_key[idx] = "TOR"
                                if "EC" in item:
                                    file_key[idx] = "EC"
                                if "EC Report" in item:
                                    file_key[idx] = "EcReport"
                                
                        # print(file_key)
                        dictionary_2 = {file_key[i]: file_link[i] for i in range(len(file_key))}
                        # print(dictionary_2)
                        dictionary_1.update(dictionary_2)
                        print(dictionary_1)
                    
                    driver.back()
                break       
                        # #
                        # df = pd.DataFrame(data=dictionary_1, index=[0])
                        # df = df.T
                        # df.to_excel("TableData.xlsx", sheet_name='Sheet1')
                