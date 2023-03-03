# Compare-sheets-google-calendar
 
This code loads several excel files with boletos from a folder, and an excel file with financial assets. It then extracts specific data from each file and creates a new dataset by merging both information based on some common columns. The final output is a list of lists containing some financial information about clients.

The function formatar_valor is used to format some monetary values in a specific Brazilian Real (BRL) currency format.

The os module is imported to access files in the file system, load_workbook from the openpyxl library is used to load and manipulate excel files, and pandas is imported as pd alias. Also, datetime and pyautogui modules are imported.

The code starts by loading all excel files with boletos from a folder, using os.path.join and a list comprehension. It then loads the excel file with financial assets using load_workbook from openpyxl and access its "Ativos" sheet.

Next, it creates three empty lists: dados_boletos, dados_ativos, and dados_individuais. Then, two for-loops are used to iterate over rows in both boletos and assets sheets and extract the desired information. These information is stored in the list dados_individuais and later added to the appropriate list (dados_boletos or dados_ativos) at each iteration.

Finally, a third for-loop is used to iterate over each list of data from dados_ativos and dados_boletos, and check if there are any matches between some columns. If a match is found, a new list with specific data from both lists is created and added to the dados list. Finally, this list is returned.

The code concludes by generating a CSV file with the data in the required format to be uploaded to Google Calendar and receive notifications of financial asset due dates. There are three tabs in the spreadsheet: one with data in free format, one where the data is organized in the spreadsheet template provided by Google, and the third tab that serves as a database. Whenever a new data is written that was not in the database, this new data is included in the tab. The CSV spreadsheet only receives new data to avoid generating duplicate entries in Google Calendar.

Overall, the code is well-structured and easy to read. The use of list comprehension and for-loops make the code efficient, and the function formatar_valor is useful to format some monetary values. The os, openpyxl, and pandas modules are well used, and the code also includes comments explaining the purpose of each section.
