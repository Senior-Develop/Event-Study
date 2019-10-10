from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.chrome.options import Options
import pandas as pd
from selenium.webdriver.support.ui import Select
import xlsxwriter



def get_driver():

    options = Options()
    options.add_experimental_option("excludeSwitches",
                                    ["ignore-certificate-errors", "safebrowsing-disable-download-protection",
                                     "safebrowsing-disable-auto-update", "disable-client-side-phishing-detection"])

    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--profile-directory=Default')
    options.add_argument("--incognito")
    options.add_argument("--disable-plugins-discovery")
    prefs = {'profile.default_content_setting_values.automatic_downloads': 1}
    options.add_experimental_option("prefs", prefs)
    #options.add_argument("--headless")
    driver = webdriver.Chrome('chromedriver', options=options)
    return driver



inputfile = 'Event.xlsx'
xlsfile = pd.ExcelFile(inputfile)

sheetnames = xlsfile.sheet_names
sheet1 = pd.read_excel(xlsfile,sheetnames[0])
sheet2 = pd.read_excel(xlsfile,sheetnames[1])
columnvl1 = sheet1["Unnamed: 0"].values
columnvl2 = sheet2["Facility event Webpage"].values
driver = get_driver()

results = []
results1 = []

for link in columnvl1:

    if "https:" in str(link):

        tbls = {}
        tbls["url"] = link
        tbls["CAT1"] = ""
        tbls["CAT2"] = ""
        tbls["CAT3"] = ""
        tbls["PELOUSE"] = ""
        tbls["CARRE OR"] = ""
        tbls["EARLY"] = ""
        tbls["FOSSE"] = ""
        tbls["DEBOUT"] = ""
        driver.get(link)
        try:
            bod = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'plan-ism')))
            pridebutton = driver.find_element_by_id("label-blocs-liens-tarifs")
            pridebutton.click()
        except:
            pass

        time.sleep(1)
        table = driver.find_element_by_id("price-table")
        rows = table.find_elements_by_tag_name("tr")
        titlearr = []
        for idx, row in enumerate(rows):
            if idx == 0:
                tds1 = row.find_elements_by_tag_name("th")
                for th_id, td1 in enumerate(tds1):
                    if th_id != 0:
                        titl = td1.text
                        titlearr.append(titl)


        for idx, title in enumerate(titlearr):
            driver.get(link)

            table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'price-table')))
            rows = table.find_elements_by_tag_name("tr")

            for idx1, row in enumerate(rows):
                if idx1 == 2:
                    tds = row.find_elements_by_tag_name("td")
                    for td_id, td in enumerate(tds):
                        if idx == td_id:
                            try:
                                Cost = Select(td.find_element_by_tag_name('select'))
                                costdrop = td.find_elements_by_tag_name("option")
                                value = len(costdrop) - 1
                                Cost.select_by_value(str(value))
                                driver.find_element_by_class_name("submitButton").click()
                                time.sleep(3)
                                try:
                                    Delete = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'action')))
                                    tbls[titlearr[td_id]] = "NOT FULL"
                                    Delete.click()
                                except:
                                    tbls[titlearr[td_id]] = "ALMOST FULL"
                                    time.sleep(3)
                            except:
                                tbls[titlearr[td_id]] = "FULL"
                    break
        results.append(tbls)



xbook = xlsxwriter.Workbook('Result.xlsx')
xsheet = xbook.add_worksheet('Facility Plan')


# write out a header row
headers = ['Facility event Webpage', 'CAT1', 'CAT2','CAT3','PELOUSE','CARRE OR','EARLY','FOSSE','DEBOUT']
bold = xbook.add_format({'bold': 1})
for i, header in enumerate(headers):
    xsheet.write(0, i, header, bold)

for row, fields in enumerate(results):
    row = row + 1
    for col, student_data in enumerate(fields):
        xsheet.write(row, col, fields[student_data])


for link in columnvl2:

    if "https:" in str(link):

        tbls = {}
        tbls["url"] = link
        tbls["ASSIS"] = ""
        tbls["DEBOUT"] = ""
        tbls["CAT1"] = ""
        tbls["CAT2"] = ""
        tbls["CAT3"] = ""
        tbls["PELOUSE"] = ""
        tbls["CARRE OR"] = ""
        tbls["DEBOUT EAR"] = ""


        driver.get(link)
        try:
            bod = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'plan-ism')))
        except:
            pass

        time.sleep(1)
        table = driver.find_element_by_id("price-table")
        rows = table.find_elements_by_tag_name("tr")
        titlearr = []
        for idx, row in enumerate(rows):
            if idx == 0:
                tds1 = row.find_elements_by_tag_name("th")
                for th_id, td1 in enumerate(tds1):
                    if th_id != 0:
                        titl = td1.text
                        titlearr.append(titl)


        for idx, title in enumerate(titlearr):
            driver.get(link)

            table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'price-table')))
            rows = table.find_elements_by_tag_name("tr")

            for idx1, row in enumerate(rows):
                if idx1 == 2:
                    tds = row.find_elements_by_tag_name("td")
                    for td_id, td in enumerate(tds):
                        if idx == td_id:
                            try:
                                Cost = Select(td.find_element_by_tag_name('select'))
                                costdrop = td.find_elements_by_tag_name("option")
                                value = len(costdrop) - 1
                                Cost.select_by_value(str(value))
                                driver.find_element_by_class_name("submitButton").click()
                                time.sleep(3)
                                try:
                                    Delete = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'action')))
                                    tbls[titlearr[td_id]] = "NOT FULL"
                                    Delete.click()
                                except:
                                    tbls[titlearr[td_id]] = "ALMOST FULL"
                                    time.sleep(3)
                            except:
                                tbls[titlearr[td_id]] = "FULL"
                    break
        results1.append(tbls)





xsheet1 = xbook.add_worksheet('Without Facilty Plan')
# write out a header row
headers = ['Facility event Webpage', 'ASSIS','DEBOUT','CAT1', 'CAT2','CAT3','PELOUSE','CARRE OR','EARLY']
bold = xbook.add_format({'bold': 1})
for i, header in enumerate(headers):
    xsheet1.write(0, i, header, bold)

for row, fields in enumerate(results1):
    row = row + 1
    for col, student_data in enumerate(fields):
        xsheet1.write(row, col, fields[student_data])

xbook.close()

driver.close()
