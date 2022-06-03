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
        /// <summary>
        /// Takes Quickbook data and adds it to a PDF file of a T4 Summary form
        /// </summary>
        /// <param name="src">Path of the T4 Summary pdf file</param>
        /// <param name="dest">Path of the filled in T4 Summary file</param>
        /// <param name="report">report class containing T4 summary attributes</param>
        /// <param name="company">company class containing information relevant to a T4 summary</param>
        public virtual void ManipulatePdf(String src, String dest, PayrollSumReport report, Company company)
        {
            //Initialize PDF document
            PdfReader pdfReader = new PdfReader(src);
            pdfReader.SetUnethicalReading(true);

            PdfDocument pdfDocument = new PdfDocument(pdfReader, new PdfWriter(dest));
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDocument, true);
            IDictionary<String, PdfFormField> fields = form.GetFormFields();
            PdfFormField toSet;

            //----Start of inputting T4 Summary

            // Date of Tax Year
            fields.TryGetValue("form1[0].Page1[0].Date[0]", out toSet);
            toSet.SetValue(report.year.Substring(2));
            // Company Name and Address
            fields.TryGetValue("form1[0].Page1[0].Border[0].EmployerInfo[0].EmployerName[0]", out toSet);
            toSet.SetValue(company.name + "\n" + company.addressFull);
            // Employee Account Number
            fields.TryGetValue("form1[0].Page1[0].Border[0].EmployerInfo[0].EmployerAccount[0]", out toSet);
            toSet.SetValue("");
            // Total number of T4 slips
            fields.TryGetValue("form1[0].Page1[0].Border[0].LeftFields[0].Line88[0].Box88[0]", out toSet);
            toSet.SetValue(report.numEmployee);
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
            // form1[0].Page1[0].Border[0].Box78[0].Box78[0]
            fields.TryGetValue("form1[0].Page1[0].Border[0].Box78[0].Box78[0]", out toSet);
            toSet.SetValue(company.phone);
            // Certification Date
            fields.TryGetValue("form1[0].Page1[0].Border[0].CertificationDate[0]", out toSet);
            toSet.SetValue(DateTime.Today.ToShortDateString().ToString());
            /*
            foreach (var item in fields)
            {
                fields.TryGetValue(item.Key, out toSet);
                toSet.SetValue(item.Key);
                Debug.WriteLine(item.Key);
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

            // Closes the PDF.
            pdfDocument.Close();
        }

        /// <summary>
        /// Defines the start of a QBXML Message Request
        /// </summary>
        /// <returns>the body of a QBXML request</returns>
        public static XmlDocument CreateXmlHeaders()
        {
            XmlDocument inputXMLDoc = new XmlDocument();
            inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
            inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"13.0\""));

            // Define Headers
            XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
            inputXMLDoc.AppendChild(qbXML);

            // Define that we want to start a request
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

            // Start parsing through the response
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

            // Save the info from response into a Report class
            report.incomeTaxDeducted = report.incomeTaxDeducted.Trim().Replace("-", String.Empty);
            report.employerCPPContribution = report.employerCPPContribution.Trim().Replace("-", String.Empty);
            report.employeeCPPContribution = report.employeeCPPContribution.Trim().Replace("-", String.Empty);
            report.employeeEIPremium = report.employeeEIPremium.Trim().Replace("-", String.Empty);
            report.employerEIPremium = report.employerEIPremium.Trim().Replace("-", String.Empty);
            report.employmentIncome = report.employmentIncome.Trim().Replace("-", String.Empty);
            report.year = year;
            
            report.calTotalDeduction();
            return report;
        }

        /// <summary>
        /// Queries the company information from Quickbook
        /// </summary>
        /// <returns>A company instance with company information filled in</returns>
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

            // Start parsing through the response
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("CompanyRet");
            XmlNodeList CompanyRet = qbXMLMsgsRsNodeList.Item(0).ChildNodes;

            Company company = new Company();

            foreach (XmlNode node in CompanyRet)
            {
                if (node.Name.Equals("CompanyName"))
                {
                   company.name = node.InnerText;
                }
                else if (node.Name.Equals("Address"))
                {
                    foreach (XmlNode innerNode in node)
                    {
                        company.addressFull += " " + innerNode.InnerText;
                    }
                    company.addressFull = company.addressFull.Trim();

                }
                else if (node.Name.Equals("AddressBlock"))
                {
                    foreach (XmlNode innerNode in node)
                    {
                        company.addressBlock += " " + innerNode.InnerText;
                    }
                    company.addressBlock = company.addressBlock.Trim();
                }
                else if (node.Name.Equals("Phone"))
                {
                    company.phone = node.InnerText;
                }
            }

/*            Console.WriteLine(company.name);
            Console.WriteLine(company.addressFull);
            Console.WriteLine(company.addressBlock);*/
            return company;
        }

        /// <summary>
        /// Grabs an existing report and adds the amount of employees paid in the given tax year
        /// </summary>
        /// <param name="report">a report instance of PayrollSumReport</param>
        /// <param name="year">specified tax year</param>
        /// <returns></returns>
        public static PayrollSumReport getEmpdata(PayrollSumReport report, String year)
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

            reportRq.AppendChild(inputXMLDoc.CreateElement("SummarizeColumnsBy")).InnerText = "Employee";

            string input = inputXMLDoc.OuterXml;

            string response = SetupConnection(input);

            // Start parsing through the response
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);

            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("NumColumns");

            // Grabs the number of columns to compute number of paid employees.
            // Paid employee has 3 columns. There are also an extra 4 rows in both the row names and the Total column.
            // Subtracting 4 and dividing it by 3 will return the number of paid employees
            report.numEmployee = ((Int16.Parse(qbXMLMsgsRsNodeList[0].InnerText) - 4) / 3).ToString() ;

            //Debug.WriteLine(response + "\n");
            //Debug.WriteLine("NumEmployee: " + report.numEmployee);

            return report;
        }
    }
}
