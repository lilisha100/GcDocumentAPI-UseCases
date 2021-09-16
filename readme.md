# GrapeCity Document API User Cases

To run the project, in Visual Studio Code teminal, run the following:

1. dotnet new console
2. dotnet add package GrapeCity.Documents.Excel
3. dotnet run

## Use Case 1

The LegacyInvoice project shows how to parse data from Excel to JSON and save the Excel to PDF.

## Use Case 2

This use case shows how to build PDF form from Excel template and then parse data from multiple PDF forms to one Excel.

The Excel2PDF project builds a PDF form from an Excel template file `Template_MemberInfo.xlsx`.

The MemberReport project parses multiple PDF forms in the `/Resources/Uploads` folder to one Excel report. The report is saved in the `Output` folder.

The `Template_MemberReport.xlsx file` under the `/Resources/Templates` folder defines how to load datasource and display the data. Once the `memberList.json` is loaded, a report will be generated per the template defined.


## Use Case 3

This LaptopServiceShop project shows how to create customized Excel from JSON and then build multiple PDFs from one Excel.
