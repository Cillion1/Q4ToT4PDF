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

        public static PayrollSumReport getPayrollSumAttribute(string response)
        {
            //Console.WriteLine(response + "\n");

            XmlDocument inputXMLDoc = CreateXmlHeaders();
            XmlElement qbXMLMsgsRq = (XmlElement)inputXMLDoc.GetElementsByTagName("QBXMLMsgsRq")[0];

            XmlElement reportRq = inputXMLDoc.CreateElement("PayrollSummaryReportQueryRq");
            qbXMLMsgsRq.AppendChild(reportRq);

            reportRq.AppendChild(inputXMLDoc.CreateElement("PayrollSummaryReportType")).InnerText = "PayrollSummary";
            reportRq.AppendChild(inputXMLDoc.CreateElement("DisplayReport")).InnerText = "false";


            XmlElement dateRange = inputXMLDoc.CreateElement("ReportPeriod");
            reportRq.AppendChild(dateRange);

            dateRange.AppendChild(inputXMLDoc.CreateElement("FromReportDate")).InnerText = "2021-01-01";
            dateRange.AppendChild(inputXMLDoc.CreateElement("ToReportDate")).InnerText = "2021-12-31";


            reportRq.AppendChild(inputXMLDoc.CreateElement("SummarizeColumnsBy")).InnerText = "TotalOnly";



            string input = inputXMLDoc.OuterXml;

            string response = SetupConnection(input);


            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);

            /*            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("PayrollSummaryReportQueryRs");
                        XmlNodeList PayrollSummaryReportQueryRs = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
                        XmlNodeList ReportRet = PayrollSummaryReportQueryRs.Item(0).ChildNodes;*/
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("ReportData");
            XmlNodeList ReportData = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
            XmlNodeList qbXMLMsgsRsNodeList2 = outputXMLDoc.GetElementsByTagName("ReportRet");
            XmlNodeList ReportRet = qbXMLMsgsRsNodeList2.Item(0).ChildNodes;

            XmlNodeList qbXMLMsgsRsNodeList3 = outputXMLDoc.GetElementsByTagName("DataRow");
            XmlNodeList DataRow = qbXMLMsgsRsNodeList3.Item(0).ChildNodes;
            XmlNodeList qbXMLMsgsRsNodeList4 = outputXMLDoc.GetElementsByTagName("SubtotalRow");
            XmlNodeList SubtotalRow = qbXMLMsgsRsNodeList4.Item(0).ChildNodes;

            string rowNumber = "";
            PayrollSumReport report = new PayrollSumReport();


            foreach (XmlNode node in ReportData)
            {
                if (node.Name.Equals("SubtotalRow"))
                {
                    foreach (XmlNode InnerNode in node)
                    {
                        if (InnerNode.Attributes["colID"].Value.Equals("1"))
                        {
                            if (InnerNode.Attributes["value"].Value.Equals("Total Gross Pay"))
                            {
                                report.employmentIncome = InnerNode.NextSibling.Attributes["value"].Value;
                            }
                        }
                    }
                }
                else if (node.Name.Equals("DataRow"))
                {
                    foreach (XmlNode InnerNode in node)
                    {
                        if (InnerNode.Attributes["value"].Value.Equals("Federal Income Tax"))
                        {
                            report.incomeTaxDeducted = InnerNode.NextSibling.Attributes["value"].Value;
                        }
                        else if (InnerNode.Attributes["value"].Value.Equals("CPP - Employee"))
                        {
                            report.employeeCPPContribution = InnerNode.NextSibling.Attributes["value"].Value;
                        }
                        else if (InnerNode.Attributes["value"].Value.Equals("EI - Employee"))
                        {
                            report.employeeEIPremium = InnerNode.NextSibling.Attributes["value"].Value;
                        }
                        else if (InnerNode.Attributes["value"].Value.Equals("CPP - Company"))
                        {
                            report.employerCPPContribution = InnerNode.NextSibling.Attributes["value"].Value;
                        }
                        else if (InnerNode.Attributes["value"].Value.Equals("EI - Company"))
                        {
                            report.employerEIPremium = InnerNode.NextSibling.Attributes["value"].Value;
                        }
                    }
                }
            }
            return report;
        }



    }
}
