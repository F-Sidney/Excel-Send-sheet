using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.Office.Interop.Outlook;


namespace ExcelSendSheet
{
    [ComVisible(true)]
    public class RibbonExt : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonExt()
        {
        }

        public void Send_Sheets_woform(Office.IRibbonControl control)
        {
            string tempfile = Path.Combine(Path.GetTempPath(), Globals.ThisWorkbook.Name.ToString());
            Globals.ThisWorkbook.SaveCopyAs(tempfile);

            try
            {
                if (ServerDocument.IsCustomized(tempfile))
                {
                    ServerDocument.RemoveCustomization(tempfile);
                }
                Microsoft.Office.Interop.Outlook.Application outlookapp = new Microsoft.Office.Interop.Outlook.Application();
                MailItem eMail = (MailItem)outlookapp.CreateItem(OlItemType.olMailItem);
                eMail.Subject = "SidneyN from elance: Workbook Attached: " + Globals.ThisWorkbook.Name.ToString();
                eMail.Attachments.Add(tempfile);
                eMail.Display(true);
                File.Delete(tempfile);
            }
            catch (System.Exception e)
            {
                //Error Removing Customization and Sending email.
            }
        }

        public void Send_Sheets_form(Office.IRibbonControl control)
        {
            Form1.ShowForm();
        }
        
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelSendSheet.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
