using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;

namespace QBToT4PDF
{

    public class InfoProcessor
    {

        public void readPDF()
        {
            string fileName = ".\\t4sum-fill-21e.pdf";

            
            PdfReader pdfRead = new PdfReader(fileName);
            PdfDocument pdfDocument = new PdfDocument(pdfRead);

            
        }

        public virtual void ManipulatePdf(String src, String dest)
        {
            //Initialize PDF document
            PdfReader pdfReader = new PdfReader(src);
            pdfReader.SetUnethicalReading(true);

            PdfDocument pdfDocument = new PdfDocument(pdfReader, new PdfWriter(dest));
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            IDictionary<String, PdfFormField> fields = form.GetFormFields();
            PdfFormField toSet;

            
            foreach (var item in fields)
            {
                fields.TryGetValue(item.Key, out toSet);
                toSet.SetValue(item.Key);
                //Debug.WriteLine(item.Key);
            }

            // form1[0].Page1[0].Border[0].EmployerInfo[0].EmployerName[0]

            //fields.TryGetValue("form1[0].Page1[0].Border[0].EmployerInfo[0].EmployerName[0]", out toSet);
            //toSet.SetValue("James Bond");

            /*
            foreach (var item in fields)
            {
                Debug.WriteLine(item.Key);
                Debug.WriteLine(item.Value);
            }
            */

            /*
            fields.TryGetValue("language", out toSet);
            toSet.SetValue("English");
            fields.TryGetValue("experience1", out toSet);
            toSet.SetValue("Off");
            fields.TryGetValue("experience2", out toSet);
            toSet.SetValue("Yes");
            fields.TryGetValue("experience3", out toSet);
            toSet.SetValue("Yes");
            fields.TryGetValue("shift", out toSet);
            toSet.SetValue("Any");
            fields.TryGetValue("info", out toSet);
            toSet.SetValue("I was 38 years old when I became an MI6 agent.");
            */
            pdfDocument.Close();
        }


    }
}
