#CREATING THE FILES AND WORKBOOK (PREFILLED)


# This section of code creates an excel file for the user to input the data she/he wants extracted and analysed
# The input file is added to the built directory and prefilled with needed information (amsterdam.02.2019)

# Create the file
import getpass
username = getpass.getuser()
import xlsxwriter 
workbook = xlsxwriter.Workbook("C:/Users/%s/Desktop/Micasaestucasa/working file/Micasa.xlsx" %username) 

# Create first worksheet, add content and format
worksheet = workbook.add_worksheet("List") 

bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'City', bold) 
worksheet.write('B1', 'Year', bold) 
worksheet.write('C1', 'Month', bold) 

# Start from the first cell. 
# Rows and columns are zero indexed. 
row = 1
column = 0
  
content = ["amsterdam","paris","barcelona", "london", "berlin"] 
  
# iterating through content list 
for item in content : 
  
    # write operation perform 
    worksheet.write(row, column, item) 
  
    # incrementing the value of row by one 
    # with each iteratons. 
    row += 1
    
row = 1
column = 1
  
content = ["2020","2019","2018", "2017"] 
  
# iterating through content list 
for item in content : 
  
    # write operation perform 
    worksheet.write(row, column, item) 
  
    # incrementing the value of row by one 
    # with each iteratons. 
    row += 1
row = 1
column = 2
  
content = ["12","11","10", "09", "08", "07", "06","05","04","03","02","01"] 
  
# iterating through content list 
for item in content : 
  
    # write operation perform 
    worksheet.write(row, column, item) 
  
    # incrementing the value of row by one 
    # with each iteratons. 
    row += 1
  
    
# Create second worksheet, add content and format. 
worksheet = workbook.add_worksheet("Instructions") 

bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'Welcome to our project, this is a little "How to" correctly utilize this file!', bold) 
worksheet.write('A2', 'Thanks to this excel file you can input the type of data that you want to extract through the online database')
worksheet.write('A3', '')
worksheet.write('A4', '1.', bold)
worksheet.write('A5', '2.', bold)
worksheet.write('A6', '3.', bold)
worksheet.write('A7', '4.', bold)

# Start from the first cell.
# Rows and columns are zero indexed.
row = 3
column = 1
content = ["The excel contains several tabs", "You can input the variables (city, year, month) in the Micasa tab","We have prefilled the Misacasa tab for ease of use, should you want to extract different data, please use information provided in the List tab",
           "Make sure that the months input until September have a ZERO (use text type)"]

# iterating through content list 
for item in content : 
  
    # write operation perform 
    worksheet.write(row, column, item) 
  
    # incrementing the value of row by one 
    # with each iteratons. 
    row += 1   

# Create third worksheet, add content and format.
worksheet = workbook.add_worksheet("Micasa") 

bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'Variable', bold) 
worksheet.write('B1', 'Input and Save the file', bold)
worksheet.write('A2', 'Month') 
worksheet.write('A3', 'Year')
worksheet.write('A4', 'City') 
 
# Start from the first cell. 
# Rows and columns are zero indexed. 
row = 1
column = 1
  
content = ["02", "2019", "amsterdam"] 
  
# iterating through content list 
for item in content : 
  
    # write operation perform 
    worksheet.write(row, column, item) 
  
    # incrementing the value of row by one 
    # with each iteratons. 
    row += 1    
workbook.close() 
 
