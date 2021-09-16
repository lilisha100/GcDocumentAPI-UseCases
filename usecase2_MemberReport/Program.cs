using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Pdf.AcroForms;

namespace MemberReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // prepare a list for billpay 
            var billPayList = new ArrayList();

            // parse every billpay pdf form data and add it into the list
            foreach (string
                inputFile
                in
                Directory
                    .EnumerateFiles(Path.Combine("Resources", "Uploads"),
                    "*.pdf")
            )
            {
                billPayList.Add(GatherData.getDataFromPdf(inputFile));
            }

            // convert the billpay list to json array
            string jsonString = JsonSerializer.Serialize(billPayList);

            // save the json file to datasource folder
            string dataSourceFile = Path.Combine("Resources","DataSource","memberList.json");
            File.WriteAllText (dataSourceFile, jsonString);

            // verify saved data
            Console.WriteLine(File.ReadAllText(dataSourceFile));

            // ------ generate excel report ----------

            // open template
            var workbook = new GrapeCity.Documents.Excel.Workbook();
            workbook.Open (Path.Combine("Resources", "Templates","Template_MemberReport.xlsx"));

            //Add data source
            var datasource =
                JsonSerializer.Deserialize<List<MemberInfos>>(File.ReadAllText(dataSourceFile));            
            workbook.AddDataSource("ds", datasource);

            //Load data by invoking to process the template
            workbook.ProcessTemplate();

            //save the report to output folder
            workbook.Save(Path.Combine("Output", "MemberReport.xlsx"));
        }
    }
}
