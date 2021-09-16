using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using GrapeCity.Documents.Excel;

namespace LegacyInvoice
{
    class Program
    {
        static void Main(string[] args)
        {
            // read worksheet from source
            Workbook workbook = new Workbook();
            workbook.Open("SimpleInvoice.xlsx");
            IWorksheet worksheet = workbook.Worksheets[0];

            // create a new invoice with blank items
            var invoice = new Invoice();
            

            // parse invoice number
            var invoiceNumberRange = worksheet.Range["F1:F1"];
            invoice.invoice_number = $"{invoiceNumberRange.Cells[0].Value}";

            // parse bill to info
            var billToRange = worksheet.Range["B5:B9"];
            int rowCount = billToRange.Rows.Count;
            String billTo = "";
            for (int i = 0; i < rowCount; i++)
            {
                billTo = billTo + billToRange.Cells[i].Value + " ";
            }
            invoice.bill_to = billTo;

            //  parse items into json array
            var items = new List<InvoiceItem>();

            var itemsRange = worksheet.Range["E7:H16"];
            int columnCount = itemsRange.Columns.Count;
            rowCount = itemsRange.Rows.Count;
            for (int i = 0; i < rowCount; i++)
            {
                if (itemsRange.Cells[i, 0].Value == null)
                {
                    break;
                }

                var item = new InvoiceItem();
                item.item_number = (string) itemsRange.Cells[i, 0].Value;
                item.description = (string) itemsRange.Cells[i, 1].Value;
                item.price = (double) itemsRange.Cells[i, 2].Value;
                item.quantity = (double) itemsRange.Cells[i, 3].Value;

                items.Add (item);
            }

            invoice.items = (List<InvoiceItem>)items;

            // save invoice to json file.
            String jsonString = JsonSerializer.Serialize(invoice);
            File.WriteAllText("invoice.json", jsonString);

            //convert invoice from excel to pdf
            workbook.Save("SimpleInvoice.pdf");
        }
    }
}
