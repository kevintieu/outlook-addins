using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RibbonDesignerXML();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Attachment_Utility___ContextMenu {
    [ComVisible(true)]
    public class RibbonDesignerXML : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;

        public RibbonDesignerXML() {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("Attachment_Utility___ContextMenu.RibbonDesignerXML.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        private FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

        public void removeAttachments_Click(Office.IRibbonControl control) {
            int itemCountTotal = 0;
            int itemCountRemoved = 0;
            int min = 0;
            Outlook.Application application = new Outlook.Application();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.Selection selected = explorer.Selection;
            DialogResult dr_Delete = MessageBox.Show("Are you sure you want to delete all attachments from the selected items?", "Warning", MessageBoxButtons.YesNo);
            if (dr_Delete == DialogResult.Yes) {
                foreach (Object obj in selected) {
                    string msg = null;
                    itemCountTotal++;
                    if (obj is Outlook.MailItem) {
                        min = 0;
                        Outlook.MailItem mailItem = (Outlook.MailItem)obj;
                        Outlook.Attachments attachments = mailItem.Attachments;
                        // Iterate through attachments of this email
                        if (attachments != null && attachments.Count > 0) {
                            int attachmentCount = attachments.Count;
                            int position = 1;
                            for (int i = 1; i <= attachmentCount; i++) {
                                Outlook.Attachment attachment = attachments[position];
                                var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");
                                String fileName = attachments[position].DisplayName;
                                // To ignore embedded attachments
                                if (flags != 4) {
                                    // To ignore embedded attachments in RTF mail with type 6
                                    if ((int)attachment.Type != 6) {
                                        // Delete attachment and append a message to the body.
                                        attachments[position].Delete();
                                        min = 1;
                                        msg += fileName + ", ";
                                    }
                                }
                                else {
                                    position++;
                                    flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");
                                }
                            }
                            if (min == 1) {
                                itemCountRemoved++;
                            }
                        }
                        if (itemCountRemoved >= 1) {
                            msg = msg.Remove(msg.Length - 2);
                            mailItem.HTMLBody = "<p>ATTACHMENTS REMOVED: [" + msg + "]</p>" + mailItem.HTMLBody;
                            mailItem.Save();
                        }
                    }
                }
                if (itemCountRemoved == 0) {
                    MessageBox.Show("There are no attachments to remove.");
                }
                else {
                    MessageBox.Show("Removed attachments from " + itemCountRemoved + " out of " + itemCountTotal + " emails.");
                }
            } else if (dr_Delete == DialogResult.No) {

            }
        }

        public void saveAttachments_Click(Office.IRibbonControl control) {
            bool firstPrompt = false;
            int itemCountTotal = 0;
            int itemCountSaved = 0;
            int validAttachments = 0;
            int min = 0;
            string path = null;
            Outlook.Application application = new Outlook.Application();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.Selection selected = explorer.Selection;
            foreach (Object obj in selected) {
                itemCountTotal++;
                if (obj is Outlook.MailItem) {
                    min = 0;
                    Outlook.MailItem mailItem = (Outlook.MailItem)obj;
                    Outlook.Attachments attachments = mailItem.Attachments;
                    if (attachments != null && attachments.Count > 0) {
                        int attachmentCount = attachments.Count;
                        // Iterate through attachments of this email
                        for (int i = 1; i <= attachmentCount; i++) {
                            Outlook.Attachment attachment = attachments[i];
                            var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");
                            String fileName = attachments[i].DisplayName;
                            // To ignore embedded attachments
                            if (flags != 4) {
                                // To ignore embedded attachments in RTF mail with type 6
                                if ((int)attachment.Type != 6) {
                                    if (firstPrompt != true) {
                                        DialogResult result = folderBrowserDialog1.ShowDialog();
                                        if (result == DialogResult.OK) {
                                            path = folderBrowserDialog1.SelectedPath;
                                        }
                                        else if (result == DialogResult.Cancel) {
                                            return;
                                        }
                                        firstPrompt = true;
                                    }
                                    validAttachments++;
                                    // If file exists, prompt user to overwrite; otherwise save the file. If user clicks yes, overwrite; user clicks no, do nothing.
                                    if (File.Exists(path + "\\" + fileName)) {
                                        DialogResult dialogResult = MessageBox.Show("A file named \"" + fileName + "\" already exists in this directory. Would you like to save this as an additional copy?", "File exists", MessageBoxButtons.YesNo);
                                        if (dialogResult == DialogResult.Yes) {
                                            int count = 1;
                                            min = 1;
                                            string fullPath = path + "\\" + fileName;
                                            string fileNameOnly = Path.GetFileNameWithoutExtension(fullPath);
                                            string extension = Path.GetExtension(fullPath);
                                            path = Path.GetDirectoryName(fullPath);
                                            string newFullPath = fullPath;
                                            while (File.Exists(newFullPath)) {
                                                string tempFileName = string.Format("{0}({1})", fileNameOnly, count++);
                                                newFullPath = Path.Combine(path, tempFileName + extension);
                                            }
                                            attachments[i].SaveAsFile(newFullPath);
                                        }
                                        else if (dialogResult == DialogResult.No) {

                                        }
                                    }
                                    else {
                                        attachments[i].SaveAsFile(path + "\\" + fileName);
                                        min = 1;
                                    }
                                }
                            }
                        }
                        if (min == 1) {
                            itemCountSaved++;
                        }
                    }
                }
            }
            if (validAttachments == 0) {
                MessageBox.Show("There are no attachments to save.");
            }
            else {
                MessageBox.Show("Saved attachments from " + itemCountSaved + " out of " + itemCountTotal + " emails.");
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
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
