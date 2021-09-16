using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;

public class AddCostInfoToServiceOrdersJson
{
    public static void doIt()
    {
        // retieve order data from json to list
        var orderList =
            JsonSerializer
                .Deserialize<List<ServiceOrder>>(File
                    .ReadAllText(Path
                        .Combine("Resources",
                        "DataSource",
                        "serviceOrders.json")));

      

        int invoice_number = 100000;
        var workbook = new GrapeCity.Documents.Excel.Workbook();

        // first, read data from hardware_cost excel file
        workbook
            .Open(Path.Combine("Uploads", "subcontract_hardware_cost.xlsx"));
        IWorksheet worksheet = workbook.Worksheets[0];

        // start from excel row 4, read service order data line by line

        for (int i = 4; i < 10; i++)
        {
            var range = worksheet.Range[$"A{i}:E{i}"];

            // exit loop when no data in first cell
            if (range.Cells[0].Value == null)
            {
                break;
            }

            // once get order_id in first cell, find corresponding invoice in the invoice list
            foreach (var item in orderList)
            {
                if ((String) range.Cells[0].Value == item.order_id)
                {
                    item.solution = (String) range.Cells[3].Value;
                    item.cost = (double) range.Cells[4].Value;
                    invoice_number = invoice_number + 1;
                    item.invoice_number = invoice_number;

                    break;
                }
            }
        }

        // second, read data from software_cost excel file then do same thing.
        workbook
            .Open(Path.Combine("Uploads", "subcontract_software_cost.xlsx"));
        worksheet = workbook.Worksheets[0];
        for (int i = 4; i < 10; i++)
        {
            var range = worksheet.Range[$"A{i}:E{i}"];
            if (range.Cells[0].Value == null)
            {
                break;
            }

            foreach (var item in orderList)
            {
                if ((String) range.Cells[0].Value == item.order_id)
                {
                    item.solution = (String) range.Cells[3].Value;
                    item.cost = (double) range.Cells[4].Value;
                    invoice_number = invoice_number + 1;
                    item.invoice_number = invoice_number;
                    break;
                }
            }
        }

        // write invoice list to json file
        File
            .WriteAllText(Path
                .Combine("Resources", "DataSource", "serviceOrders.json"),
            JsonSerializer.Serialize(orderList));
    }
}
