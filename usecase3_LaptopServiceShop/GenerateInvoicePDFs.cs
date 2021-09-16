using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;

public class GenerateInvoicePDFs
{
    public static void doIt()
    {
        var workbook = new GrapeCity.Documents.Excel.Workbook();

        // read json file into array
        var invoiceList =
            JsonSerializer
                .Deserialize<List<ServiceOrder>>(File
                    .ReadAllText(Path
                        .Combine("Resources", "DataSource", "serviceOrders.json")));


        foreach (var invoice in invoiceList)
        {
            workbook
                .Open(Path
                    .Combine("Resources",
                    "Templates",
                    "Template_Invoice.xlsx"));

            workbook.AddDataSource("ds", invoice);

            workbook.ProcessTemplate();

            // add order_id as sufix of invoice file name
            workbook
                .Save(Path
                    .Combine("Downloads", $"Invoice_{invoice.order_id}.pdf"));
        }
    }
}
