using System.Collections;
using System.Collections.Generic;
public class Invoice
{
    public string invoice_number { get; set; }

    public string bill_to { get; set; }

    public List<InvoiceItem> items {get; set;}

}
