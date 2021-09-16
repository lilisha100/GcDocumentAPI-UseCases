using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;

public class GenerateCostExcel
{
    public static void doIt()
    {
        // first, we extrace all orders and put into software and hardware list.
        var orderList =
            JsonSerializer
                .Deserialize<List<ServiceOrder>>(File
                    .ReadAllText(Path
                        .Combine("Resources",
                        "DataSource",
                        "serviceOrders.json")));
        var softwareOrderList = new ArrayList();
        var hardwareOrderList = new ArrayList();

        foreach (var item in orderList)
        {
            if (item.service_type == "hardware")
            {
                hardwareOrderList.Add (item);
            }
            else if (item.service_type == "software")
            {
                softwareOrderList.Add (item);
            }
        }

        // then load the list data to template then export them
        
        var workbook = new GrapeCity.Documents.Excel.Workbook();

        // open template
        workbook
            .Open(Path
                .Combine("Resources",
                "Templates",
                "Template_Service_Cost.xlsx"));

        //  bind a source
        workbook.AddDataSource("ds", softwareOrderList);

        // load data to template
        workbook.ProcessTemplate();

        //export cost excel for subcontract to download
        workbook
            .Save(Path.Combine("Downloads", "subcontract_software_cost.xlsx"));

        // do the same thing on hardware order data
        workbook
            .Open(Path
                .Combine("Resources",
                "Templates",
                "Template_Service_Cost.xlsx"));

        workbook.AddDataSource("ds", hardwareOrderList);

        workbook.ProcessTemplate();

        workbook
            .Save(Path.Combine("Downloads", "subcontract_hardware_cost.xlsx"));
    }
}
