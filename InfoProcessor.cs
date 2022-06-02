using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Interop.QBXMLRP2;
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

        public virtual void ManipulatePdf(String src, String dest, PayrollSumReport report)
        {
            //Initialize PDF document
            PdfReader pdfReader = new PdfReader(src);
            pdfReader.SetUnethicalReading(true);

            PdfDocument pdfDocument = new PdfDocument(pdfReader, new PdfWriter(dest));
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            IDictionary<String, PdfFormField> fields = form.GetFormFields();
            PdfFormField toSet;


            //----Start of inputting T4 Summary
            // form1[0].Page1[0].Border[0].EmployerInfo[0].Employe]rName[0
            //Box 14 - Emploment Income
            fields.TryGetValue("form1[0].Page1[0].Border[0].LeftFields[0].Line14[0].Box14[0]", out toSet);
            toSet.SetValue(report.employmentIncome);
            //Box 16 - Employees CPP Contribution
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line16[0].Box16[0]", out toSet);
            toSet.SetValue(report.employeeCPPContribution);
            //Box 27 - Employer's CPP contributions
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line27[0].Box27[0]", out toSet);
            toSet.SetValue(report.employerCPPContribution);
            //Box 18 - Employees' EI premiums
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line18[0].Box18[0]", out toSet);
            toSet.SetValue(report.employeeEIPremium);
            //Box 19 - Employer's EI premiums
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line19[0].Box19[0]", out toSet);
            toSet.SetValue(report.employerEIPremium);
            //Box 22 - Income tax deducted
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line22[0].Box22[0]", out toSet);
            toSet.SetValue(report.incomeTaxDeducted);
            //Box 80 - Total deductions reported
            fields.TryGetValue("form1[0].Page1[0].Border[0].MiddleFields[0].Line80[0].Box80[0]", out toSet);
            toSet.SetValue(report.totalDeductionsReported);


            /*
            foreach (var item in fields)
            {
                fields.TryGetValue(item.Key, out toSet);
                toSet.SetValue(item.Key);
                //Debug.WriteLine(item.Key);
            }
            */


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

        public static XmlDocument CreateXmlHeaders()
        {
            XmlDocument inputXMLDoc = new XmlDocument();
            inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
            inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

            // Headers
            XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
            inputXMLDoc.AppendChild(qbXML);

            XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");

            return inputXMLDoc;
        }


        /// <summary>
        /// Sends the given XML string to Quickbook and retrieves their response
        /// </summary>
        /// <param name="input">a XML string </param>
        /// <returns>the XML response string from Quickbook</returns>
        public static string SetupConnection(string input)
        {
            RequestProcessor2 rp = null;
            string ticket = null;
            string response = null;
            try
            {
                rp = new RequestProcessor2();
                rp.OpenConnection("", "IDN EmployeeAdd C# sample");
                ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
                response = rp.ProcessRequest(ticket, input);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                //MessageBox.Show("COM Error Description = " + ex.Message, "COM error");
                return ex.Message;
            }
            finally
            {
                if (ticket != null)
                {
                    rp.EndSession(ticket);
                }
                if (rp != null)
                {
                    rp.CloseConnection();
                }
            };
            return response;
        }

        /// <summary>
        /// Getting data for T4 Summary Report
        /// </summary>
        /// <param name="year">the year of the T4 Summary report </param>
        /// <returns> an object that holds data for the T4 Summary Report Quickbook</returns>
        public static PayrollSumReport getPayrollSumAttribute(string year)
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

            dateRange.AppendChild(inputXMLDoc.CreateElement("FromReportDate")).InnerText = year + "-01-01";
            dateRange.AppendChild(inputXMLDoc.CreateElement("ToReportDate")).InnerText = year + "-12-31";

            reportRq.AppendChild(inputXMLDoc.CreateElement("SummarizeColumnsBy")).InnerText = "TotalOnly";

            string input = inputXMLDoc.OuterXml;

            string response = SetupConnection(input);


            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);

            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("ReportData");
            XmlNodeList ReportData = qbXMLMsgsRsNodeList.Item(0).ChildNodes;

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

            report.incomeTaxDeducted = report.incomeTaxDeducted.Trim().Replace("-", String.Empty);
            report.employerCPPContribution = report.employerCPPContribution.Trim().Replace("-", String.Empty);
            report.employeeCPPContribution = report.employeeCPPContribution.Trim().Replace("-", String.Empty);
            report.employeeEIPremium = report.employeeEIPremium.Trim().Replace("-", String.Empty);
            report.employerEIPremium = report.employerEIPremium.Trim().Replace("-", String.Empty);
            report.employmentIncome = report.employmentIncome.Trim().Replace("-", String.Empty);
            report.calTotalDeduction();
            return report;
        }


        public static Company getCompanyInfo()
        {

            XmlDocument inputXMLDoc = CreateXmlHeaders();
            XmlElement qbXMLMsgsRq = (XmlElement)inputXMLDoc.GetElementsByTagName("QBXMLMsgsRq")[0];

            XmlElement CompanyQueryRq = inputXMLDoc.CreateElement("CompanyQueryRq");
            qbXMLMsgsRq.AppendChild(CompanyQueryRq);

            //CompanyQueryRq.AppendChild(inputXMLDoc.CreateElement("IncludeRetElement")).InnerText = "EIN";

            string input = inputXMLDoc.OuterXml;

            string response = SetupConnection(input);

            Console.WriteLine(response);

            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);

            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("CompanyRet");
            XmlNodeList CompanyRet = qbXMLMsgsRsNodeList.Item(0).ChildNodes;

            Company company = new Company();

            foreach (XmlNode node in CompanyRet)
            {
                if (node.Name.Equals("CompanyName"))
                    company.name = node.InnerText;
                else if (node.Name.Equals("Address"))
                {
                    foreach (XmlNode innerNode in node)
                    {
                        company.addressFull += " " + innerNode.InnerText;
                    }

                }
                else if (node.Name.Equals("AddressBlock"))
                {
                    foreach (XmlNode innerNode in node)
                    {
                        company.addressBlock += " " + innerNode.InnerText;
                    }

                }

            }

/*            Console.WriteLine(company.name);
            Console.WriteLine(company.addressFull);
            Console.WriteLine(company.addressBlock);*/
            return company;
        }

    }
}
