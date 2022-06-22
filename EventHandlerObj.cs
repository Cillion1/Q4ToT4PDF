/*
 * Class that handles an event when the menu item is clicked.
 * 
 * For most development purposes, you only need to modify parts under UIExtensionEvent.
 */

using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;  // For using MessageBox.
using QBSDKEVENTLib; // In order to implement IQBEventCallback.
using System.Runtime.InteropServices;  // For use of the GuidAttribute, ProgIdAttribute and ClassInterfaceAttribute.
using System.Xml; //XML Parsing
using System.IO;

namespace QBToT4PDF
{
    [
      Guid("62447F81-C195-446f-8201-94F0614E49D5"),  // We indicate a specific CLSID for "QBToT4PDF.EventHandlerObj" for convenience of searching the registry.
      ProgId("SubscribeAndHandleQBEvent.EventHandlerObj"),  // This ProgId is used by default. Not 100% necessary.
      ClassInterface(ClassInterfaceType.None)
    ]
    public class EventHandlerObj :
        ReferenceCountedObjectBase, // EventHandlerObj is derived from ReferenceCountedObjectBase so that we can track its creation and destruction.
        IQBEventCallback  // this must implement the IQBEventCallback interface.
    {

        public EventHandlerObj()
        {
            // ReferenceCountedObjectBase constructor will be invoked.
            Console.WriteLine("EventHandlerObj constructor.");
        }

        ~EventHandlerObj()
        {
            // ReferenceCountedObjectBase destructor will be invoked.
            Console.WriteLine("EventHandlerObj destructor.");
        }

        //Call back function which would be invoked from the QB
        public void inform(string strMessage)
        {
            try
            {
                StringBuilder sb = new StringBuilder(strMessage);
                XmlDocument outputXMLDoc = new XmlDocument();
                outputXMLDoc.LoadXml(strMessage);
                XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("QBXMLEvents");
                XmlNode childNode = qbXMLMsgsRsNodeList.Item(0).FirstChild;

                // handle the event based on type of event
                switch (childNode.Name)
                {
                    case "DataEvent":
                        //Handle Data Event Here
                        MessageBox.Show(sb.ToString(), "DATA EVENT - From QB");
                        break;

                    case "UIExtensionEvent":
                        // Run our functions here when we click the menu item.
                        OpenT4Form();
                        
                        break;

                    default:
                        MessageBox.Show(sb.ToString(), "Response From QB");
                        break;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error in processing the response from QB - " + ex.Message);
            }
        }

        /// <summary>
        /// Function to run when an event is executed
        /// 
        /// Opens a file dialog to grab the T4 pdf file and creates a new file with a filled-in T4 form file.
        /// 
        /// NOTE: must have Quickbook with a company file open
        /// </summary>
        public static void OpenT4Form()
        {
            // Year to grab report from. If you want the current year at all times. Use DateTime.Now.Year
            string year = "2021";

            // Define class instances to store information from Quickbook
            PayrollSumReport report = InfoProcessor.GetPayrollSumAttribute(year);
            report = InfoProcessor.GetEmployeeData(report, year);
            Company company = InfoProcessor.GetCompanyInfo();

            // Find the T4 summary page and create a new copy of it with a filled form.
            try
            {
                // Grabs the directory of the pdf file
                string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Application.ExecutablePath)).ToString();

                // Default T4 file name. File from https://www.canada.ca/en/revenue-agency/services/forms-publications/forms/t4.html
                string fileName = "t4sum-fill-21e";

                // Find the T4 pdf file
                string src = filePath + "\\" + fileName + ".pdf";

                // Location and name of filled pdf. This should be the same directory as the original filled pdf.
                string endDest = filePath + "\\" + fileName + " - Filled.pdf";

                Console.WriteLine("Creating filled T4 Summary at " + endDest);

                // Create base filled pdf
                FileInfo file = new FileInfo(endDest);
                file.Directory.Create();

                // Start filling the T4 Summary pdf.
                new InfoProcessor().FillT4PDF(src, endDest, report, company);

                MessageBox.Show("Finished Creating T4 PDF file");
            } catch (Exception ex)
            {
                // Error message if something wrong happens.
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }

    class EventHandlerObjClassFactory : ClassFactoryBase
	{
		public override void virtual_CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
		{
            Console.WriteLine("EventHandlerObjClassFactory.CreateInstance().");
			Console.WriteLine("Requesting Interface : " + riid.ToString());

            if (riid == Marshal.GenerateGuidForType(typeof(IQBEventCallback)) ||
                riid == Program.IID_IDispatch ||
                riid == Program.IID_IUnknown)
			{
                EventHandlerObj EventHandlerObj_New = new EventHandlerObj();

                ppvObject = Marshal.GetComInterfaceForObject(EventHandlerObj_New, typeof(IQBEventCallback));
			} 
			else
			{
				throw new COMException("No interface",  unchecked((int) 0x80004002));
			}
		}
	}
}