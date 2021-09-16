using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;

namespace LaptopServiceShop
{
    /**
--- Business Scenario  ---

Our Laptop Service Shop provides services such as parts repair, system recover... 

When clients send their laptops to us, we outsource sorfware service orders to one subcontract and hareware to the other one. 

we need the subcontracts download cost excel files to fill what have been done on the orders and how much they cost.  

Once the cost excel files are filled and uploaded back, we generate invoice pdfs for customer to download.

*/
    class Program
    {
        static void Main(string[] args)
        {
            // generate some sample service orders. See serviceOrders.json under Resources/DataSource/
            GenerateSampleJson.doIt();

            // a template Template_Service_Cost.xlsx in under Resources/Templates folder. We will load the sample orders to the template and export the result in Downloads/ folder.
            GenerateCostExcel.doIt();

            // pretending we are the subcontracts. now we open the cost excel files in the download folder, fill solution and cost data, then move the files to Uploads/ folder.
            // Once the filled excels are in place, we parse the date and update the serviceOrder.json file with it
            AddCostInfoToServiceOrdersJson.doIt();

            // a template Template_Invlice.xlsx is under Resources/Templates folder. We generate invoice PDF files based on the template and invoice data.
            // the pdf invoice files are exported under Downloads folder for customers to download.
            GenerateInvoicePDFs.doIt();
        }
    }
}
