from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from easygui import *
import win32com.client as win32
import time

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chrome_options=options)

msg = "Enter shortcode login information"
title = "Holder Analysis automation login"
fieldNames = ["Shortcode Email Address", "Password"]
fieldValues = []  # we start with blanks for the values
fieldValues = multpasswordbox(msg, title, fieldNames)

# make sure that none of the fields was left blank

while 1:
    if fieldValues is None: break
    errmsg = ""
    for i in range(len(fieldNames)):
        if fieldValues[i].strip() == "":
            errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
    if errmsg == "": break  # no problems found
    fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)

driver.get("https://adfs.hscic.gov.uk/adfs/ls/?client-request-id=1c3c0d01-f7f3-4be9-a491-ec618036a275&wa=wsignin1.0&wtrealm=urn%3afederation%3aMicrosoftOnline&wctx=LoginOptions%3D3%26estsredirect%3d2%26estsrequest%3drQIIAZVSO2zTYBCOm7a0FYWqYmBg6NAyIDn9_Sd-JKISfaRJ09huHKfGXirHj_h34kcc2yGeOjIWgYTEWJg6woJY6NyprB0ZEK8BwQAjjlBZmPh0-nR3Op0-3Xd3skSOKC2DP6DxMf8h7TK7RLA4t4DN37314-en3RePvqqH1HF0ggErDP1BaXW157k43mkHWuD0UGwMOhQo5NKCIHL6yFUdpA1ymuesvsawcwz7iGEnEwMqTxcAkS8WqSIgmUKBLOTYrTLBOpzNQTZkkw6UEQBp7dSlHShLjZAVW0AR9x1WFGzWEWyloiDF3u7ylUao2A2CS-c52IJ10XJ4SQ4VUbA4pwV4USb5ynbvYuI6vx6FFhyTF6DE-D4xa3qBc-B7g_BZ9nSC9w13R9_0XNfQwtx4zHBDpKkh8ty9wPONIETGYI1dT7GjjdSgVUNWNYgOVBXRsjHyKqre6FoRuzcy0HCn2OS6TbpZiNebu4Hk0qHdr-X7CunG5FaxUx5v2RjTZpMcKZLga86-rW8Knl4VhlrixXVoJTLs2WwiE3WHG7XF8oO6I5BtZztUJS5Je3G7ydwWDB0FqWLR-1faCjT_iquT-2Y7rgTahkb3e1uS4VA8lXQ1RyjXysZ6r8EPE83so2AoWUCvbiQ8YmK1ytmqxCac3YrkfC11oRwp1V4kS4Qv59lIhsWwPlzJ62m8zC6TwKQATZh4u20aeGqxijMMyOMUnRrOGBQsGvAsey29pYv0JT_wTNQz3mcXA7WtMvesgYa0XMeLc1H3dPLw-dT5JPZ58irIlmZm5hYyNzNLmV-T2PFU-o9flj6823qMVZ9k3wr001eZs6lVNybcAsPRA9K_L3obNT-S-NC0g5gxY3W_D2OKTTsS8qnGGlUijqaxo-npb9PYwyuZN7P__c0XczcggAAnCByQSwCWIFUq5JXT-cxv0&cbcxt=&mkt=&lc=")
driver.find_element_by_id("userNameInput").send_keys(fieldValues[0])
driver.find_element_by_id("passwordInput").send_keys(fieldValues[1])
driver.find_element_by_id("submitButton").click()

wait = WebDriverWait(driver, 20)

bypass_staying_signed_in = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@value="No"]')))
driver.find_element_by_xpath('//input[@value="No"]').click()

driver.get("https://hscic365.crm11.dynamics.com/crmreports/viewer/viewer.aspx?action=filter&id=9d3b28aa-ad14-e911-a9e2-000d3a2bb31e&helpID=DARS-ProcessStage_v5new.rdl")
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'resultFrame')))
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_1_89iT0_aria')]/div[1]")))

DARS_Summary_OpenByStage_Total = driver.find_element_by_xpath("//div[contains(@id, '_1_89iT0_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Subtotal = driver.find_element_by_xpath("//div[contains(@id, '_1_71iT0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Subtotal_Breach = driver.find_element_by_xpath("//div[contains(@id, '1_67iT0R0x0C0x1_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Total_Breach = driver.find_element_by_xpath("//div[contains(@id, '_1_85iT0C0x1_aria')]/div[1]").get_attribute('innerHTML')

driver.find_element_by_id("reportViewer_ctl08_ctl04_ctl01").click()
driver.find_element_by_id("reportViewer_ctl08_ctl04_divDropDown_ctl00").click()
driver.find_element_by_id("reportViewer_ctl08_ctl04_divDropDown_ctl10").click()
driver.find_element_by_id("reportViewer_ctl08_ctl00").click()
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_2_89iT0_aria')]/div[1]")))

DARS_Summary_OpenByStage_Total_CCG = driver.find_element_by_xpath( "//div[contains(@id, '_2_89iT0_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Subtotal_CCG = driver.find_element_by_xpath( "//div[contains(@id, '_2_71iT0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Subtotal_Breach_CCG = driver.find_element_by_xpath("//div[contains(@id, '2_67iT0R0x0C0x1_aria')]/div[1]").get_attribute('innerHTML')
DARS_Summary_OpenByStage_Total_Breach_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_85iT0C0x1_aria')]/div[1]").get_attribute('innerHTML')

driver.get("https://hscic365.crm11.dynamics.com/crmreports/viewer/viewer.aspx?action=filter&id=2242ce07-bac0-e811-a9d8-000d3a2bbda1&helpID=DARS-OpenAndClosuresWeeklySummary_v3new.rdl")
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'resultFrame')))
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_1_71iT1R0R0R0x0_aria')]/div[1]")))

DARS_OpenClosures_Triage_Failures = driver.find_element_by_xpath( "//div[contains(@id, '_1_71iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Applications_Accepted = driver.find_element_by_xpath( "//div[contains(@id, '_1_75iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Signed_DSA = driver.find_element_by_xpath( "//div[contains(@id, '_1_79iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Submissions = driver.find_element_by_xpath( "//div[contains(@id, '_1_83iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')

driver.find_element_by_id("reportViewer_ctl08_ctl08_ctl01").click()
driver.find_element_by_id("reportViewer_ctl08_ctl08_divDropDown_ctl00").click()
driver.find_element_by_id("reportViewer_ctl08_ctl08_divDropDown_ctl10").click()
driver.find_element_by_id("reportViewer_ctl08_ctl00").click()
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_2_71iT1R0R0R0x0_aria')]/div[1]")))

DARS_OpenClosures_Triage_Failures_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_71iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Applications_Accepted_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_75iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Signed_DSA_CCG = driver.find_element_by_xpath( "//div[contains(@id, '_2_79iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Submissions_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_83iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')

driver.get("https://hscic365.crm11.dynamics.com/crmreports/viewer/viewer.aspx?action=filter&id=f380b8bd-c6b1-e811-a9d6-000d3a2bb91c&helpID=DARS-AverageSLA_v3new.rdl")
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'resultFrame')))
wait.until(EC.presence_of_element_located((By.ID, "reportViewer_ctl08_ctl04_ctl01")))
driver.find_element_by_id("reportViewer_ctl08_ctl04_ctl01").click()
driver.find_element_by_id("reportViewer_ctl08_ctl04_divDropDown_ctl00").click()
driver.find_element_by_id("reportViewer_ctl08_ctl00").click()

wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_1_55iT0_aria')]/div[1]")))
Combined_Mean_Working_Days = driver.find_element_by_xpath("//div[contains(@id, '_1_55iT0_aria')]/div[1]").get_attribute(
    'innerHTML')

driver.find_element_by_id("reportViewer_ctl08_ctl04_ctl01").click()
driver.find_element_by_id("reportViewer_ctl08_ctl04_divDropDown_ctl00").click()
driver.find_element_by_id("reportViewer_ctl08_ctl04_divDropDown_ctl10").click()
driver.find_element_by_id("reportViewer_ctl08_ctl00").click()

wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_2_55iT0_aria')]/div[1]")))
Combined_Mean_Working_Days_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_55iT0_aria')]/div[1]").get_attribute('innerHTML')

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=4c951b2a-a566-e911-a98a-00224800c940&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction1 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=d330fc0f-986a-e911-a98c-00224800cf35&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction2 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=8744b1b6-8f6a-e911-a98c-00224800cf35&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction3 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=a2a4f834-3668-e911-a988-00224800c719&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction4 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=0102d6bf-7822-ea11-a810-000d3a86d801&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction5 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=queueitem&viewid=34d9b8fc-84c8-e811-a9dd-000d3a2bb190&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
VH_DARS_Queue_Items = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=8d222296-fc7d-e911-a98a-00224800c940&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[2]/span")))
time.sleep(3)
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span").text).split()
HolderAnalysis_DSA_Financial_YTD = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=4f474203-bfd0-e911-a813-000d3a86d6fd&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_DSA_Financial_YTD_NewDSA = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=account&viewid=4ccb5865-b7d3-e911-a813-000d3a86d68d&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span").text).split()
HolderAnalysis_Org_DSA_Signed_Financial_YTD = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_applicationholderdsa&viewid=09252713-8811-e811-8128-70106fa55dc1&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
Outstanding_Triage = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_applicationholderdsa&viewid=220d137d-ba33-ea11-a813-000d3a86d535&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
Outstanding_Triage_CCG = found_string[4]


driver.get("https://app.powerbi.com/groups/7e8fcf98 -1b8e-47e4-a10d-4bbd0e9f425c/reports/9c089b54-2d0d-4d15-9e2e-bdb4f3d86b97/ReportSectiond29fad801f74e3f6bf8c")

wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                     "-content/div/div/report/exploration-container/exploration"
                                                     "-container-modern/div/div/div/exploration-host/div/div"
                                                     "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                     "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                     "2]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                     "1]/div[6]/div/div/div[1]/div")))
time.sleep(3)

Email_Tracked_To_Holder = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                       "-content/div/div/report/exploration-container/exploration"
                                                       "-container-modern/div/div/div/exploration-host/div/div"
                                                       "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                       "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                       "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                       "1]/div[6]/div/div/div[1]/div").text

Average_Age_of_Email = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content"
                                                    "/div/div/report/exploration-container/exploration-container"
                                                    "-modern/div/div/div/exploration-host/div/div/exploration/div"
                                                    "/explore-canvas-modern/div/div[2]/div/div[2]/div["
                                                    "2]/visual-container-repeat/visual-container-modern["
                                                    "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                    "1]/div[6]/div/div/div[2]/div").text

Average_Age_of_Data_Application_Email = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container"
                                                                     "/mat-sidenav-content/div/div/report/exploration"
                                                                     "-container/exploration-container-modern/div/div"
                                                                     "/div/exploration-host/div/div/exploration/div"
                                                                     "/explore-canvas-modern/div/div[2]/div/div["
                                                                     "2]/div["
                                                                     "2]/visual-container-repeat/visual-container"
                                                                     "-modern[4]/transform/div/div["
                                                                     "3]/div/visual-modern/div/div/div[2]/div[1]/div["
                                                                     "4]/div/div/div[2]/div[1]").text

Data_Application_Email_Count = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                            "-content/div/div/report/exploration-container/exploration"
                                                            "-container-modern/div/div/div/exploration-host/div/div"
                                                            "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                            "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                            "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                            "1]/div[4]/div/div/div[1]/div[1]").text

Average_Age_of_Data_Production_Email = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat"
                                                                    "-sidenav-content/div/div/report/exploration"
                                                                    "-container/exploration-container-modern/div/div"
                                                                    "/div/exploration-host/div/div/exploration/div"
                                                                    "/explore-canvas-modern/div/div[2]/div/div["
                                                                    "2]/div["
                                                                    "2]/visual-container-repeat/visual-container"
                                                                    "-modern[4]/transform/div/div["
                                                                    "3]/div/visual-modern/div/div/div[2]/div[1]/div["
                                                                    "4]/div/div/div[2]/div[2]").text

Data_Production_Email_Count = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                           "-content/div/div/report/exploration-container/exploration"
                                                           "-container-modern/div/div/div/exploration-host/div/div"
                                                           "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                           "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                           "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                           "1]/div[4]/div/div/div[1]/div[2]").text

Not_Attached_To_Holder_Count = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                            "-content/div/div/report/exploration-container"
                                                            "/exploration-container-modern/div/div/div/exploration"
                                                            "-host/div/div/exploration/div/explore-canvas-modern/div"
                                                            "/div[2]/div/div[2]/div["
                                                            "2]/visual-container-repeat/visual-container-modern["
                                                            "2]/transform/div/div[3]/div/visual-modern/div/div/div["
                                                            "2]/div[1]/div[6]/div/div/div[1]/div").text

Average_Age_Not_Attached_To_Holder = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat"
                                                                  "-sidenav-content/div/div/report/exploration"
                                                                  "-container/exploration-container-modern/div/div"
                                                                  "/div/exploration-host/div/div/exploration/div"
                                                                  "/explore-canvas-modern/div/div[2]/div/div[2]/div["
                                                                  "2]/visual-container-repeat/visual-container"
                                                                  "-modern[2]/transform/div/div["
                                                                  "3]/div/visual-modern/div/div/div[2]/div[1]/div["
                                                                  "6]/div/div/div[2]/div").text

driver.get("https://app.powerbi.com/groups/7e6fa73a-fc03-421c-8de9-e405f86dc62f/reports/53ef3e82-3680-457e-9027-7942c75dca2a/ReportSection")
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.card')))
time.sleep(3)
found_string = str(driver.find_element_by_css_selector('.card').text).split()
Open_at_1c = found_string[0]
#add count column to the report to make it easier to grab this value
#add an index of all holders in the report starting at 1 with an increment of 0 then add total

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_file&viewid=b536c135-f481-e911-a98d-00224800bb9b&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span")))
driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[1]/div/div/ul/li[8]/div/ul/li/button").click()
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[7]/div/div/ul/li[1]/ul/li[1]/button")))
driver.find_element_by_xpath("/html/body/div[7]/div/div/ul/li[1]/ul/li[1]/button").click()
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'wopi_frame')))
wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'ewrch-row-cellsel')]")))
driver.find_element_by_xpath("//*[contains(@class, 'ewrch-row-cellsel')]").click()
actions = ActionChains(driver)
actions.key_down(Keys.LEFT_CONTROL)
actions.send_keys(Keys.ARROW_DOWN)
actions.key_up(Keys.SHIFT)
actions.perform()
wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'ewrch-row-nosel ewrch-row-pre-cellsel')]")))

HolderAnalysis_DataDisseminationsFinancialYTD = driver.find_element_by_xpath("//*[contains(@class, 'ewrch-row-nosel ewrch-row-pre-cellsel')]").text

driver.quit()

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

file = 'C:/Users/Rashad/Documents/Python Projects/Automate_Friday/Test_Book.xlsx'
wb = excel.Workbooks.Open(file)
ws = wb.Worksheets('Total_Apps_AutoFill')

ws.Range("A16:A16").EntireRow.Insert()
ws.Range("A18:A17").AutoFill(ws.Range("A18:A16"), win32.constants.xlFillDefault)
ws.Range("E18:E17").AutoFill(ws.Range("E18:E16"), win32.constants.xlFillDefault)
ws.Range("F18:F17").AutoFill(ws.Range("F18:F16"), win32.constants.xlFillDefault)
ws.Range("Q18:Q17").AutoFill(ws.Range("Q18:Q16"), win32.constants.xlFillDefault)
ws.Range("R18:R17").AutoFill(ws.Range("R18:R16"), win32.constants.xlFillDefault)
ws.Range("U18:U17").AutoFill(ws.Range("U18:U16"), win32.constants.xlFillDefault)
ws.Range("X18:X17").AutoFill(ws.Range("X18:X16"), win32.constants.xlFillDefault)
ws.Cells(16, 2).Value = DARS_Summary_OpenByStage_Total
ws.Cells(16, 3).Value = DARS_Summary_OpenByStage_Subtotal
ws.Cells(16, 4).Value = DARS_Summary_OpenByStage_Subtotal_Breach
ws.Cells(16, 7).Value = DARS_Summary_OpenByStage_Total_Breach
ws.Cells(16, 8).Value = Email_Tracked_To_Holder
ws.Cells(16, 9).Value = Outstanding_Triage
ws.Cells(16, 10).Value = round(float(Average_Age_of_Email))
ws.Cells(16, 11).Value = DARS_OpenClosures_Triage_Failures
ws.Cells(16, 12).Value = DARS_OpenClosures_Applications_Accepted
ws.Cells(16, 13).Value = DARS_OpenClosures_Signed_DSA
ws.Cells(16, 14).Value = DARS_OpenClosures_Submissions
ws.Cells(16, 15).Value = Open_at_1c
ws.Cells(16, 19).Value = round(float(Combined_Mean_Working_Days))
ws.Cells(16, 22).Value = HolderAnalysis_Data_Destruction1
ws.Cells(16, 23).Value = HolderAnalysis_Data_Destruction2
ws.Cells(16, 25).Value = HolderAnalysis_Data_Destruction5
ws.Cells(16, 26).Value = HolderAnalysis_Data_Destruction3
ws.Cells(16, 27).Value = HolderAnalysis_Data_Destruction4
ws.Cells(16, 28).Value = round(float(Average_Age_of_Data_Application_Email))
ws.Cells(16, 29).Value = Data_Application_Email_Count
ws.Cells(16, 30).Value = round(float(Average_Age_of_Data_Production_Email))
ws.Cells(16, 31).Value = Data_Production_Email_Count

ws = wb.Worksheets('Total_Apps_CCG_AutoFill')

ws.Range("A13:A13").EntireRow.Insert()
ws.Range("A15:A14").AutoFill(ws.Range("A15:A13"), win32.constants.xlFillDefault)
ws.Range("M15:M14").AutoFill(ws.Range("M15:M13"), win32.constants.xlFillDefault)
ws.Range("S15:S14").AutoFill(ws.Range("S15:S13"), win32.constants.xlFillDefault)
ws.Cells(13, 2).Value = DARS_Summary_OpenByStage_Total_CCG
ws.Cells(13, 3).Value = DARS_Summary_OpenByStage_Subtotal_CCG
ws.Cells(13, 4).Value = DARS_Summary_OpenByStage_Subtotal_Breach_CCG
ws.Cells(13, 5).Value = DARS_Summary_OpenByStage_Total_Breach_CCG
ws.Cells(13, 6).Value = VH_DARS_Queue_Items
ws.Cells(13, 7).Value = Outstanding_Triage_CCG
ws.Cells(13, 9).Value = DARS_OpenClosures_Triage_Failures_CCG
ws.Cells(13, 10).Value = DARS_OpenClosures_Applications_Accepted_CCG
ws.Cells(13, 11).Value = DARS_OpenClosures_Signed_DSA_CCG
ws.Cells(13, 12).Value = DARS_OpenClosures_Submissions_CCG
ws.Cells(13, 14).Value = round(float(Combined_Mean_Working_Days_CCG))

ws = wb.Worksheets('Enq_NotAttached_AutoFill')

ws.Range("A44:A44").EntireRow.Insert()
ws.Range("A46:A45").AutoFill(ws.Range("A46:A44"), win32.constants.xlFillDefault)
ws.Cells(44, 2).Value = Not_Attached_To_Holder_Count
ws.Cells(44, 3).Value = round(float(Average_Age_Not_Attached_To_Holder))

ws = wb.Worksheets('YTD_Numbers_AutoFill')

ws.Cells(13, 4).Value = HolderAnalysis_DSA_Financial_YTD
ws.Cells(14, 4).Value = HolderAnalysis_DSA_Financial_YTD_NewDSA
ws.Cells(15, 4).Value = HolderAnalysis_Org_DSA_Signed_Financial_YTD
ws.Cells(17, 8).Value = HolderAnalysis_DataDisseminationsFinancialYTD

ws = wb.Worksheets('Total_Apps_Formula_AutoFill')

ws.Range("A13:A13").EntireRow.Insert()
ws.Range("A15:O14").AutoFill(ws.Range("A15:O13"), win32.constants.xlFillDefault)

print("yum")