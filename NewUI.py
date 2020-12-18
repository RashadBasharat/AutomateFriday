from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import easygui as eg
import win32com.client as win32
import time

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(chrome_options=options)
# driver = Chrome()

# username = eg.enterbox("Please enter your shortcode email address")
# password = eg.enterbox("Please enter your password")

driver.get("https://adfs.hscic.gov.uk/adfs/ls/?client-request-id=1c3c0d01-f7f3-4be9-a491-ec618036a275&wa=wsignin1.0&wtrealm=urn%3afederation%3aMicrosoftOnline&wctx=LoginOptions%3D3%26estsredirect%3d2%26estsrequest%3drQIIAZVSO2zTYBCOm7a0FYWqYmBg6NAyIDn9_Sd-JKISfaRJ09huHKfGXirHj_h34kcc2yGeOjIWgYTEWJg6woJY6NyprB0ZEK8BwQAjjlBZmPh0-nR3Op0-3Xd3skSOKC2DP6DxMf8h7TK7RLA4t4DN37314-en3RePvqqH1HF0ggErDP1BaXW157k43mkHWuD0UGwMOhQo5NKCIHL6yFUdpA1ymuesvsawcwz7iGEnEwMqTxcAkS8WqSIgmUKBLOTYrTLBOpzNQTZkkw6UEQBp7dSlHShLjZAVW0AR9x1WFGzWEWyloiDF3u7ylUao2A2CS-c52IJ10XJ4SQ4VUbA4pwV4USb5ynbvYuI6vx6FFhyTF6DE-D4xa3qBc-B7g_BZ9nSC9w13R9_0XNfQwtx4zHBDpKkh8ty9wPONIETGYI1dT7GjjdSgVUNWNYgOVBXRsjHyKqre6FoRuzcy0HCn2OS6TbpZiNebu4Hk0qHdr-X7CunG5FaxUx5v2RjTZpMcKZLga86-rW8Knl4VhlrixXVoJTLs2WwiE3WHG7XF8oO6I5BtZztUJS5Je3G7ydwWDB0FqWLR-1faCjT_iquT-2Y7rgTahkb3e1uS4VA8lXQ1RyjXysZ6r8EPE83so2AoWUCvbiQ8YmK1ytmqxCac3YrkfC11oRwp1V4kS4Qv59lIhsWwPlzJ62m8zC6TwKQATZh4u20aeGqxijMMyOMUnRrOGBQsGvAsey29pYv0JT_wTNQz3mcXA7WtMvesgYa0XMeLc1H3dPLw-dT5JPZ58irIlmZm5hYyNzNLmV-T2PFU-o9flj6823qMVZ9k3wr001eZs6lVNybcAsPRA9K_L3obNT-S-NC0g5gxY3W_D2OKTTsS8qnGGlUijqaxo-npb9PYwyuZN7P__c0XczcggAAnCByQSwCWIFUq5JXT-cxv0&cbcxt=&mkt=&lc=")
driver.find_element_by_id("userNameInput").send_keys("raba8@hscic.gov.uk")
driver.find_element_by_id("passwordInput").send_keys("Muffin786")
driver.find_element_by_id("submitButton").click()

wait = WebDriverWait(driver, 15)

annoying_login = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@value="No"]')))
driver.find_element_by_xpath('//input[@value="No"]').click()


driver.get("https://hscic365.crm11.dynamics.com/crmreports/viewer/viewer.aspx?action=filter&id=2242ce07-bac0-e811-a9d8-000d3a2bbda1&helpID=DARS-OpenAndClosuresWeeklySummary_v3new.rdl")
wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'resultFrame')))
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_1_71iT1R0R0R0x0_aria')]/div[1]")))

DARS_OpenClosures_Triage_Failures = driver.find_element_by_xpath("//div[contains(@id, '_1_71iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Applications_Accepted = driver.find_element_by_xpath("//div[contains(@id, '_1_75iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Signed_DSA = driver.find_element_by_xpath("//div[contains(@id, '_1_79iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Submissions = driver.find_element_by_xpath("//div[contains(@id, '_1_83iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')

print("openclosures")
print(DARS_OpenClosures_Triage_Failures)
print(DARS_OpenClosures_Applications_Accepted)
print(DARS_OpenClosures_Signed_DSA)
print(DARS_OpenClosures_Submissions)

driver.find_element_by_id("reportViewer_ctl08_ctl08_ctl01").click()
driver.find_element_by_id("reportViewer_ctl08_ctl08_divDropDown_ctl00").click()
driver.find_element_by_id("reportViewer_ctl08_ctl08_divDropDown_ctl10").click()
driver.find_element_by_id("reportViewer_ctl08_ctl00").click()
wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@id, '_2_71iT1R0R0R0x0_aria')]/div[1]")))

DARS_OpenClosures_Triage_Failures_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_71iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Applications_Accepted_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_75iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Signed_DSA_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_79iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')
DARS_OpenClosures_Submissions_CCG = driver.find_element_by_xpath("//div[contains(@id, '_2_83iT1R0R0R0x0_aria')]/div[1]").get_attribute('innerHTML')

print("openclosureccg")
print(DARS_OpenClosures_Triage_Failures_CCG)
print(DARS_OpenClosures_Applications_Accepted_CCG)
print(DARS_OpenClosures_Signed_DSA_CCG)
print(DARS_OpenClosures_Submissions_CCG)

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=4c951b2a-a566-e911-a98a-00224800c940&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction1 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=d330fc0f-986a-e911-a98c-00224800cf35&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction2 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=8744b1b6-8f6a-e911-a98c-00224800cf35&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction3 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=a2a4f834-3668-e911-a988-00224800c719&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction4 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=0102d6bf-7822-ea11-a810-000d3a86d801&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_Data_Destruction5 = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=queueitem&viewid=34d9b8fc-84c8-e811-a9dd-000d3a2bb190&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
VH_DARS_Queue_Items = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=8d222296-fc7d-e911-a98a-00224800c940&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[2]/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span").text).split()
HolderAnalysis_DSA_Financial_YTD = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_application&viewid=4f474203-bfd0-e911-a813-000d3a86d6fd&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
HolderAnalysis_DSA_Financial_YTD_NewDSA = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=account&viewid=4ccb5865-b7d3-e911-a813-000d3a86d68d&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span").text).split()
HolderAnalysis_Org_DSA_Signed_Financial_YTD = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_applicationholderdsa&viewid=09252713-8811-e811-8128-70106fa55dc1&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
Outstanding_Triage = found_string[4]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_applicationholderdsa&viewid=220d137d-ba33-ea11-a813-000d3a86d535&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span")))
found_string = str(driver.find_element_by_xpath("/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div/span").text).split()
Outstanding_Triage_CCG = found_string[4]

print("dashboard stuff")
print(HolderAnalysis_Data_Destruction1)
print(HolderAnalysis_Data_Destruction2)
print(HolderAnalysis_Data_Destruction3)
print(HolderAnalysis_Data_Destruction4)
print(HolderAnalysis_Data_Destruction5)
print(VH_DARS_Queue_Items)
print(HolderAnalysis_DSA_Financial_YTD)
print(HolderAnalysis_DSA_Financial_YTD_NewDSA)
print(HolderAnalysis_Org_DSA_Signed_Financial_YTD)
print(Outstanding_Triage)
print(Outstanding_Triage_CCG)


driver.get("https://app.powerbi.com/groups/7e8fcf98-1b8e-47e4-a10d-4bbd0e9f425c/reports/9c089b54-2d0d-4d15-9e2e-bdb4f3d86b97/ReportSectiond29fad801f74e3f6bf8c")

wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                     "-content/div/div/report/exploration-container/exploration"
                                                     "-container-modern/div/div/div/exploration-host/div/div"
                                                     "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                     "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                     "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                     "1]/div[6]/div/div/div[1]/div")))

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

Data_Application_Email = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
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

Data_Production_Email = driver.find_element_by_xpath("/html/body/div[1]/root/mat-sidenav-container/mat-sidenav"
                                                      "-content/div/div/report/exploration-container/exploration"
                                                      "-container-modern/div/div/div/exploration-host/div/div"
                                                      "/exploration/div/explore-canvas-modern/div/div[2]/div/div["
                                                      "2]/div[2]/visual-container-repeat/visual-container-modern["
                                                      "4]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div["
                                                      "1]/div[4]/div/div/div[1]/div[2]").text

print("Queue Stats")
print(Email_Tracked_To_Holder)
print(Average_Age_of_Email)
print(Average_Age_of_Data_Application_Email)
print(Data_Application_Email)
print(Average_Age_of_Data_Production_Email)
print(Data_Production_Email)


driver.get("https://app.powerbi.com/groups/7e6fa73a-fc03-421c-8de9-e405f86dc62f/reports/53ef3e82-3680-457e-9027-7942c75dca2a/ReportSection")
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.card')))
found_string = str(driver.find_element_by_css_selector('.card').text).split()
Open_at_1c = found_string[0]

driver.get("https://hscic365.crm11.dynamics.com/main.aspx?app=d365default&forceUCI=1&pagetype=entitylist&etn=cps_file&viewid=b536c135-f481-e911-a98d-00224800bb9b&viewType=4230")
wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div[4]/div[2]/div/div/div/div/div[2]/div/section/div[3]/div/div[2]/div[2]/div/div[1]/span")))
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

print("app at 1c and the 5000 number")
print(Open_at_1c)
print(HolderAnalysis_DataDisseminationsFinancialYTD)

driver.quit()







# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = True
#
# file = 'C:/Users/Rashad/Documents/Python Projects/Automate_Friday/Test_Book.xlsx'
# wb = excel.Workbooks.Open(file)
# ws = wb.Worksheets('Test')
# print(ws.Name)
#
# ws.Range("A16:A16").EntireRow.Insert()
# ws.Range("A18:A17").AutoFill(ws.Range("A18:A16"), win32.constants.xlFillDefault)
# ws.Range("E18:E17").AutoFill(ws.Range("E18:E16"), win32.constants.xlFillDefault)
# ws.Range("F18:F17").AutoFill(ws.Range("F18:F16"), win32.constants.xlFillDefault)
# ws.Range("Q18:Q17").AutoFill(ws.Range("Q18:Q16"), win32.constants.xlFillDefault)
# ws.Range("R18:R17").AutoFill(ws.Range("R18:R16"), win32.constants.xlFillDefault)
# ws.Range("U18:U17").AutoFill(ws.Range("U18:U16"), win32.constants.xlFillDefault)
# ws.Range("X18:X17").AutoFill(ws.Range("X18:X16"), win32.constants.xlFillDefault)
# ws.Cells(16, 11).Value = 5
# ws.Cells(16, 12).Value = 27
# ws.Cells(16, 13).Value = 20
# ws.Cells(16, 14).Value = 35