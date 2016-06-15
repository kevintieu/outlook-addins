using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Save_and_Send {

    public partial class RibbonDesigner {

        private void RibbonDesigner_Load(object sender, RibbonUIEventArgs e) {
            Console.WriteLine("Save and Send Loaded");
        }

        private void saveSend_Click(object sender, RibbonControlEventArgs e) {
            Outlook.Application application = new Outlook.Application();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook._MailItem mailItem = inspector.CurrentItem;
            Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
            Outlook.Folder folder = (Outlook.Folder)nameSpace.PickFolder();
            mailItem.SaveSentMessageFolder = folder;
            mailItem.GetInspector.Activate();
            System.Windows.Forms.SendKeys.SendWait("%S");
        }
    }
}
