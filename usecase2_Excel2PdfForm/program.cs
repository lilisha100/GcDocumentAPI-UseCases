using Newtonsoft.Json;
using System;
using System.Collections;
using System.Numerics;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Expressions;
using GrapeCity.Documents.Excel.Drawing;
using System.Globalization;

namespace GrapeCity.Documents.Excel.Examples.Templates.PDFFormBuilderSamples.CustomerBillPay
{
    class Program
    {
        static void Main(string[] args)
        {
		//create a new workbook
		var workbook = new GrapeCity.Documents.Excel.Workbook();
		
		//Load template file from resource
		workbook.Open("Template_MemberInfo.xlsx");
		
		//No data needed
		
		//Invoke to process the template
		workbook.ProcessTemplate();
		        
		//save to a pdf file
		workbook.Save("MemberInfo.pdf");

        }

    }
}