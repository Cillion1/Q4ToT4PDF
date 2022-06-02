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
                        //Handle UI Extension Event HERE
                        //MessageBox.Show(sb.ToString(), "UI EXTENSION EVENT - From QB. Start running Code");
                        //Application.Run(new mainDashboardUI());

                        // Need to full path the T4 Form
                        // TODO: Add a way to select a directory in Windows filesystems
                        //string fileName = ".\\t4sum-fill-21e.pdf";

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

        public static void OpenT4Form()
        {
            PayrollSumReport report = InfoProcessor.getPayrollSumAttribute("2021");
            report = InfoProcessor.getEmpdata(report, "2021");
            Company company = InfoProcessor.getCompanyInfo();


            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "pdf",
                Filter = "pdf files (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                Debug.WriteLine(filePath);

                FileInfo file_info = new FileInfo(filePath);
                string fileName = Path.GetFileNameWithoutExtension(file_info.ToString());
                Debug.WriteLine(fileName);

                string endDest = file_info.DirectoryName + "\\" + fileName + "123.pdf";
                Debug.WriteLine(endDest);

                FileInfo file = new FileInfo(endDest);
                file.Directory.Create();

                //InfoProcessor processor = new InfoProcessor();
                new InfoProcessor().ManipulatePdf(filePath, endDest, report, company);
                MessageBox.Show("Finished Creating T4 PDF file", "T4 Form");

            }
            else
            {
                MessageBox.Show("Error: Could not fill in T4 form");
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