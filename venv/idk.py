excel = win32.Dispatch('Excel.Application')  # Opens up excel
# excel.Visible = True  # Makes excel visible, this can be changed to false if you dont want it to pop up

file = 'C:/Users/Rashad/Documents/Python Projects/Automate_Friday/Test_Book.xlsx'  # Location of the automation book
wb = excel.Workbooks.Open(file)
ws = wb.Worksheets('Total_Apps_AutoFill')  # After opening the file as wb (workbook) you now go to individual worksheets

ws.Range("A16:A16").EntireRow.Insert()  # Goes to A16 and just inserts a new row
ws.Range("A18:A17").AutoFill(ws.Range("A18:A16"), win32.constants.xlFillDefault)  # Formula so just pulled up from 2 cells below as autofill
ws.Range("E18:E17").AutoFill(ws.Range("E18:E16"), win32.constants.xlFillDefault)
ws.Range("F18:F17").AutoFill(ws.Range("F18:F16"), win32.constants.xlFillDefault)
ws.Range("Q18:Q17").AutoFill(ws.Range("Q18:Q16"), win32.constants.xlFillDefault)
ws.Range("R18:R17").AutoFill(ws.Range("R18:R16"), win32.constants.xlFillDefault)
ws.Range("U18:U17").AutoFill(ws.Range("U18:U16"), win32.constants.xlFillDefault)
ws.Range("X18:X17").AutoFill(ws.Range("X18:X16"), win32.constants.xlFillDefault)

ws.Cells(16, 2).Value = DARS_Summary_OpenByStage_Total  # Using the direct Y, X coordinate to fill the cell with the value
ws.Cells(16, 3).Value = DARS_Summary_OpenByStage_Subtotal
ws.Cells(16, 4).Value = DARS_Summary_OpenByStage_Subtotal_Breach
ws.Cells(16, 7).Value = DARS_Summary_OpenByStage_Total_Breach
ws.Cells(16, 8).Value = Email_Tracked_To_Holder
ws.Cells(16, 9).Value = Outstanding_Triage
ws.Cells(16, 10).Value = round(float(Average_Age_of_Email))  # These are typecasted to float and rounded as needed
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

ws = wb.Worksheets('Total_Apps_CCG_AutoFill')  # Again switching worksheet and repeating

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

wb.Save()  # Saves workbook
wb.Close()  # Closes workbook, can be commented out if you want to have a look
excel.Application.Quit()

print("yum")
