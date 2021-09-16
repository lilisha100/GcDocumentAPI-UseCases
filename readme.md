# read invoice data from excel and write it to json obj
# convert excel to pdf

# the demo only processs a simgle file. There are other demos showing how to process multiple files.

--- init ---
dotnet new console
dotnet add package GrapeCity.Documents.Excel

--- get a simple Excel file from the demo, and put it under project folder ---
https://www.grapecity.com/documents-api-excel/demos/simpleinvoice

--- copy content in Program.cs ---

--- dotnet run ---

# --- part 1 ---
# The customerbillpay.pdf file under Resources/Templates folder is a blank PDF form for clients to fill.
# Filled pdf forms are placed in Resources/Uploads folder. 
# Main program reads data from the forms and converts it to json file as datasource.
# --- part 2 ---
# The Template_BillPayReport.xlsx file under Resources/Templates folder defines how to load datasource and display the data.
# Once the billPayList.json is loaded, a report will generated per the template defined.
# The BillPayReort is saved in the Output folder

--- init ---
dotnet new console
dotnet add package GrapeCity.Documents.Excel

--- create file structure ---
Resources/Templates folder should have template files
Resources/DataSource folder leaves blank. a json file will be saved here after paring data.
Output folder leaves blank. the report will be generated here.

--- copy BillPayInfos.cs to root folder ---
It is a class definination, defines attributtes in BillPayInfos object

--- copy GatherData.cs to root folder ---
it is used as a utility, which contains a function return a billPay object. the input is a pdf filename.

