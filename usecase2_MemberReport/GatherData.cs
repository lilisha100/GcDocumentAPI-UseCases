using System;
using System.IO;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Pdf.AcroForms;

public class GatherData
{
    // read one pdf form and save to a member object
    public static MemberInfos getDataFromPdf(string pdfFileName)
    {
        GcPdfDocument doc = new GcPdfDocument();

        // The original file stream must be kept open while working with the loaded PDF, see LoadPDF for details:
        using (
            var fs =
                new FileStream(pdfFileName,
                    FileMode.Open,
                    FileAccess.Read)
        )
        {
            doc.Load (fs);
            var page = doc.Pages.Last;
            var member = new MemberInfos();
            foreach (Field field in doc.AcroForm.Fields)
            {
                switch (field.Name)
                {
                    case "first_name":
                        member.first_name = (String) field.Value;
                        break;
                    case "last_name":
                        member.last_name = (String) field.Value;
                        break;
                    case "address":
                        member.address = (String) field.Value;
                        break;
                    case "address_2":
                        member.address_2 = (String) field.Value;
                        break;
                    case "city":
                        member.city = (String) field.Value;
                        break;
                    case "state":
                        member.state = (String) field.Value;
                        break;
                    case "postal_code":
                        member.postal_code = (String) field.Value;
                        break;
                    case "phone_area":
                        member.phone_area = (String) field.Value;
                        break;
                    case "phone_number":
                        member.phone_number = (String) field.Value;
                        break;
                    case "email":
                        member.email = (String) field.Value;
                        break;
                    case "member_id":
                        member.member_id = (string) field.Value;
                        break;
                    case "fee_due":
                        member.fee_due = (string) field.Value;
                        break;
                    default:
                        break;
                }
            }

            return member;
        }
    }
}
