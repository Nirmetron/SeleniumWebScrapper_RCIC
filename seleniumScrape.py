import logging
import time
import traceback

import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

driver = webdriver.Chrome()
driver.implicitly_wait(15)
workbook = xlsxwriter.Workbook("RCIC Data.xlsx")
wrap_format = workbook.add_format({'text_wrap': True})
worksheet = workbook.add_worksheet("RCIC DATA")
order = 1
link = ""
failsafe = 2
person = []
all_links = []
spreadsheet_labels = ["Order", "Name", "College ID", "Eligible for Service", "Current Licence", "Recent Status Change",
                      "Licence Status",
                      "Licence Class", "Licence Start Date", "Licence Expiry Date", "Licence Status",
                      "Suspension Status", "Suspension Reason", "Suspension Start Date", "Suspension End Date",
                      "Suspension End Reason",
                      "Employment Companies", "Start Date", "Country", "Province", "City", "Email", "Phone"]
# there was also , "Agents Name", "Agents Company", "Agents Start Date", "Agents Country", "Agents Province", "Agents City", "Agents Phone", "Agents Email"
# but they were removed because agents table is fake
wrap_format.set_font_name("Times New Roman")
wrap_format.set_font_size(14)
wrap_format.set_align('center')
wrap_format.set_align('vcenter')
for col_num, data in enumerate(spreadsheet_labels):
    worksheet.write(0, col_num, data, wrap_format)
driver.get("https://register.college-ic.ca/Public-Register-EN/RCIC_Search.aspx")
original_window = driver.current_window_handle


def get_person_links():
    global links
    try:
        while links[0] == driver.find_element(By.LINK_TEXT, "Select"):
            time.sleep(0.05)
        links = driver.find_elements(By.LINK_TEXT, "Select")
        for elem_link in links:
            # print(link.get_attribute("href"))
            all_links.append(elem_link.get_attribute("href"))
            # get_person_info(link)
        driver.find_element(By.CSS_SELECTOR, ".rgPageNext").click()
    # time.sleep(2)
    except Exception as e:
        logging.error(traceback.format_exc())
        print("error in get_links " + str(e))


def get_person_info(person_link):
    global order, person, wrap_format, failsafe
    try:
        person.clear()
        person.append(order)
        driver.get(person_link)
        # driver.find_element(By.LINK_TEXT, "Licensee Details").click()
        # name
        person.append(WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".card-body p:nth-of-type(1) span"))).text)
        # ID
        person.append(WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR,
                                          "div#ctl01_TemplateBody_WebPartManager1_gwpciPersonDetails_ciPersonDetails__Body>div>section>div>div>div"))).text)
        # Eligible
        person.append(WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section:nth-of-type(1) strong span"))).text)
        # Current Licence
        person.append(WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR,
                                          "div#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_currentLicn__Body>div>section>div>div>div:nth-child(1)>span:nth-child(2)"))).text)
        # Date of Licence Change
        person.append(driver.find_element(By.CSS_SELECTOR,
                                          "div#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_currentLicn__Body>div>section>div>div>div:nth-child(2)>span:nth-child(3)").text)
        # Licence Status
        person.append(driver.find_element(By.CSS_SELECTOR,
                                          "div#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_currentLicn__Body>div>section>div>div>div:nth-child(3)>span:nth-child(2)").text)
        # Licence Class
        licence_class_elem = WebDriverWait(driver, 25).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR,
                                                  "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_LicenceHistory_ResultsGrid_Grid1_ctl00 td:nth-of-type(1)")))
        # Check if licence info is present
        if licence_class_elem[0].text != "There are no records.":
            licence_classes = ""
            for elem_data in licence_class_elem:
                licence_classes += elem_data.text + "\n"
            licence_classes = licence_classes[:-1]
            person.append(licence_classes)
            # Licence Start Date
            licence_start_elem = driver.find_elements(By.CSS_SELECTOR,
                                                      "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_LicenceHistory_ResultsGrid_Grid1_ctl00 td:nth-of-type(2)")
            licence_start = ""
            for elem_data in licence_start_elem:
                licence_start += elem_data.text + "\n"
            licence_start = licence_start[:-1]
            person.append(licence_start)
            # Licence Expiry Date
            licence_expiry_elem = driver.find_elements(By.CSS_SELECTOR,
                                                       "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_LicenceHistory_ResultsGrid_Grid1_ctl00 td:nth-of-type(3)")
            licence_expiry = ""
            for elem_data in licence_expiry_elem:
                licence_expiry += elem_data.text + "\n"
            licence_expiry = licence_expiry[:-1]
            person.append(licence_expiry)
            # Licence Status
            licence_status_elem = driver.find_elements(By.CSS_SELECTOR,
                                                       "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_LicenceHistory_ResultsGrid_Grid1_ctl00 td:nth-of-type(4)")
            licence_status = ""
            for elem_data in licence_status_elem:
                licence_status += elem_data.text + "\n"
            licence_status = licence_status[:-1]
            person.append(licence_status)
        else:
            while len(person) != 11:
                person.append("No Record")
        # Suspension Status
        suspension_status_elem = driver.find_elements(By.CSS_SELECTOR,
                                                      "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_SuspensionRevocation_ResultsGrid_Grid1_ctl00 td:nth-of-type(1)")
        # Check if suspension info is present
        if suspension_status_elem[0].text != "There are no records.":
            suspension_status = ""
            for elem_data in suspension_status_elem:
                suspension_status += elem_data.text + "\n"
            suspension_status = suspension_status[:-1]
            person.append(suspension_status)
            # Suspension Reason
            suspension_reason_elem = driver.find_elements(By.CSS_SELECTOR,
                                                          "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_SuspensionRevocation_ResultsGrid_Grid1_ctl00 td:nth-of-type(2)")
            suspension_reason = ""
            for elem_data in suspension_reason_elem:
                suspension_reason += elem_data.text + "\n"
            suspension_reason = suspension_reason[:-1]
            person.append(suspension_reason)
            # Suspension Start Date
            suspension_start_elem = driver.find_elements(By.CSS_SELECTOR,
                                                         "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_SuspensionRevocation_ResultsGrid_Grid1_ctl00 td:nth-of-type(3)")
            suspension_start = ""
            for elem_data in suspension_start_elem:
                suspension_start += elem_data.text + "\n"
            suspension_start = suspension_start[:-1]
            person.append(suspension_start)
            # Suspension End Date
            suspension_end_elem = driver.find_elements(By.CSS_SELECTOR,
                                                       "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_SuspensionRevocation_ResultsGrid_Grid1_ctl00 td:nth-of-type(4)")
            suspension_end = ""
            for elem_data in suspension_end_elem:
                suspension_end += elem_data.text + "\n"
            suspension_end = suspension_end[:-1]
            person.append(suspension_end)
            # Suspension End Reason
            suspension_end_reason_elem = driver.find_elements(By.CSS_SELECTOR,
                                                              "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_SuspensionRevocation_ResultsGrid_Grid1_ctl00 td:nth-of-type(5)")
            suspension_end_reason = ""
            for elem_data in suspension_end_reason_elem:
                suspension_end_reason += elem_data.text + "\n"
            suspension_end_reason = suspension_end_reason[:-1]
            person.append(suspension_end_reason)
        else:
            while len(person) != 16:
                person.append("No Record")
        # Employment Company
        employment_company_elem = driver.find_elements(By.CSS_SELECTOR,
                                                       "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(1)")
        # Check for employment info
        if employment_company_elem[0].text != "There are no records.":
            employment_company = ""
            for elem_data in employment_company_elem:
                employment_company += elem_data.text + "\n"
            employment_company = employment_company[:-1]
            person.append(employment_company)
            # Employment Start Date
            employment_start_elem = driver.find_elements(By.CSS_SELECTOR,
                                                         "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(2)")
            employment_start = ""
            for elem_data in employment_start_elem:
                employment_start += elem_data.text + "\n"
            employment_start = employment_start[:-1]
            person.append(employment_start)
            # Employment Country
            employment_country_elem = driver.find_elements(By.CSS_SELECTOR,
                                                           "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(3)")
            employment_country = ""
            for elem_data in employment_country_elem:
                employment_country += elem_data.text + "\n"
            employment_country = employment_country[:-1]
            person.append(employment_country)
            # Employment Province
            employment_province_elem = driver.find_elements(By.CSS_SELECTOR,
                                                            "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(4)")
            employment_province = ""
            for elem_data in employment_province_elem:
                if elem_data.text != employment_province[:-1]:
                    employment_province += elem_data.text + "\n"
            employment_province = employment_province[:-1]
            person.append(employment_province)
            # Employment City
            employment_city_elem = driver.find_elements(By.CSS_SELECTOR,
                                                        "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(5)")
            employment_city = ""
            for elem_data in employment_city_elem:
                if elem_data.text != employment_city[:-1]:
                    employment_city += elem_data.text + "\n"
            employment_city = employment_city[:-1]
            person.append(employment_city)
            # Email
            employment_email_elem = driver.find_elements(By.CSS_SELECTOR,
                                                         "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(6)")
            employment_email = ""
            for elem_data in employment_email_elem:
                if elem_data.text != employment_email[:-1]:
                    employment_email += elem_data.text + "\n"
            employment_email = employment_email[:-1]
            person.append(employment_email)
            # Phone
            employment_company_elem = driver.find_elements(By.CSS_SELECTOR,
                                                           "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Employment_ResultsGrid_Grid1_ctl00 td:nth-of-type(7)")
            employment_company = ""
            for elem_data in employment_company_elem:
                if elem_data.text != employment_company[:-1]:
                    employment_company += elem_data.text + "\n"
            employment_company = employment_company[:-1]
            person.append(employment_company)
        else:
            while len(person) != 23:
                person.append("No Record")
        # ACTUALLY AGENTS INFO IS FAKE AS OF KNOW
        # UNCOMMENT WHEN IT BECOMES REAL
        # PART OF THE LABELS FOR AGENTS WILL ALSO BE REMOVED
        # # Agents Name
        # This below is the real agent_name_elem that will work when the agents table will become an actual table.
        # Uncomment this one and remove the current one when agents become a real thing
        # agent_name_elem = driver.find_elements(By.CSS_SELECTOR, "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(1)")
        agent_name_elem = driver.find_elements(By.CSS_SELECTOR,
                                               "#ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td")
        # Check of agent info is present
        if agent_name_elem[0].text != "There are no records.":
            print("HOLY SHIT AGENTS ARE REAL")
            print(person)
            # agent_name = ""
            # for elem_data in agent_name_elem:
            #     agent_name += elem_data.text + "\n"
            # agent_name = agent_name[:-1]
            # person.append(agent_name)
            # # Agents Company
            # agent_company_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                           "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(2)")
            # agent_company = ""
            # for elem_data in agent_company_elem:
            #     agent_company += elem_data.text + "\n"
            # agent_company = agent_company[:-1]
            # person.append(agent_company)
            # # Agents Start Date
            # agent_start_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                         "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(3)")
            # agent_start = ""
            # for elem_data in agent_start_elem:
            #     agent_start += elem_data.text + "\n"
            # agent_start = agent_start[:-1]
            # person.append(agent_start)
            # # Agents Country
            # agent_country_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                           "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(4)")
            # agent_country = ""
            # for elem_data in agent_country_elem:
            #     agent_country += elem_data.text + "\n"
            # agent_country = agent_country[:-1]
            # person.append(agent_country)
            # # Agents Province
            # agent_province_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                            "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(5)")
            # agent_province = ""
            # for elem_data in agent_province_elem:
            #     agent_province += elem_data.text + "\n"
            # agent_province = agent_province[:-1]
            # person.append(agent_province)
            # # Agents City
            # agent_city_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                        "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(6)")
            # agent_city = ""
            # for elem_data in agent_city_elem:
            #     agent_city += elem_data.text + "\n"
            # agent_city = agent_city[:-1]
            # person.append(agent_city)
            # # Agents Phone
            # agent_phone_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                         "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(7)")
            # agent_phone = ""
            # for elem_data in agent_phone_elem:
            #     if elem_data.text != agent_phone[:-1]:
            #         agent_phone += elem_data.text + "\n"
            # agent_phone = agent_phone[:-1]
            # person.append(agent_phone)
            # # Agents Email
            # agent_email_elem = driver.find_elements(By.CSS_SELECTOR,
            #                                         "ctl01_TemplateBody_WebPartManager1_gwpciProfileCCO_ciProfileCCO_Agents_ResultsGrid_Grid1_ctl00 td:nth-of-type(8)")
            # agent_email = ""
            # for elem_data in agent_email_elem:
            #     if elem_data.text != agent_email[:-1]:
            #         agent_email += elem_data.text + "\n"
            # agent_email = agent_email[:-1]
            # person.append(agent_email)
        # else:
        #     while len(person) != 31:
        #         person.append("No Record")
    except Exception as e:
        if failsafe:
            print("Retry ", failsafe, "/2")
            failsafe -= 1
            get_person_info(link)
        logging.error(traceback.format_exc())
        print(link)
    failsafe = 2
    for column, p_data in enumerate(person):
        worksheet.write(order, column, p_data, wrap_format)
    order += 1


driver.find_element(By.CLASS_NAME, "TextButton").click()
btn_last = driver.find_element(By.CLASS_NAME, "rgPageLast")
btn_last.click()
time.sleep(4)
lastpage_num = int(driver.find_element(By.CLASS_NAME, "rgCurrentPage").text)
# Debug print last page num
print(lastpage_num, "last page")
driver.find_element(By.CLASS_NAME, "rgPageFirst").click()
time.sleep(4)
links = [driver.find_element(By.CSS_SELECTOR,
                             "#ctl01_TemplateBody_WebPartManager1_gwpciSearchLicensee_ciSearchLicensee_ResultsGrid_Grid1_ctl00__5 a")]
for x in range(lastpage_num):
# for debuging
# for debuging
# for x in range(10):
    get_person_links()

while len(all_links) != 0:
    # &b9100e1006f6=2#b9100e1006f6 just open licence details directly
    check_link = all_links.pop() + "&b9100e1006f6=2#b9100e1006f6"
    if check_link == link:
        print("Person duplicate skipped ", check_link)
        continue
    else:
        link = check_link
    get_person_info(link)

# close everything
worksheet.autofit()
workbook.close()
driver.close()
