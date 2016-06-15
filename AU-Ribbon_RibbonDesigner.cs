using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;


namespace Attachment_Utility___Ribbon {
    public partial class RibbonDesigner {

        private void RibbonDesigner_Load(object sender, RibbonUIEventArgs e) {
            Console.WriteLine("Ribbon loaded.");
        }

        private void RemoveAttachment_Click(object sender, RibbonControlEventArgs e) {
            int itemCountTotal = 0;
            int itemCountRemoved = 0;
            int min = 0;
            var explorer = (Outlook.Explorer)e.Control.Context;
            Outlook.Selection selected = explorer.Selection;
            DialogResult dr_delete = MessageBox.Show("Are you sure you want to delete all attachments from the selected items?", "Warning", MessageBoxButtons.YesNo);
            if (dr_delete == DialogResult.Yes) {
                foreach (Object obj in selected) {
                    string msg = null;
                    itemCountTotal++;
                    if (obj is Outlook.MailItem) {
                        min = 0;
                        Outlook.MailItem mailItem = (Outlook.MailItem)obj;
                        Outlook.Attachments attachments = mailItem.Attachments;
                        if (attachments != null && attachments.Count > 0) {
                            int attachmentCount = attachments.Count;
                            int position = 1;
                            // Iterate through the attachments of this email
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
                            if (msg != null) {
                                msg = msg.Remove(msg.Length - 2);
                                mailItem.HTMLBody = "<p>ATTACHMENTS REMOVED: [" + msg + "]</p>" + mailItem.HTMLBody;
                                mailItem.Save();
                            }
                        }
                    }
                }
                if (itemCountRemoved == 0) {
                    MessageBox.Show("There are no attachments to remove.");
                }
                else {
                    MessageBox.Show("Removed attachments from " + itemCountRemoved + " out of " + itemCountTotal + " emails.");
                }
            } else if (dr_delete == DialogResult.No) {

            }
        }

        private void SaveAttachment_Click(object sender, RibbonControlEventArgs e) {
            bool firstPrompt = false;
            int itemCountTotal = 0;
            int itemCountSaved = 0;
            int validAttachments = 0;
            int min = 0;
            var explorer = (Microsoft.Office.Interop.Outlook.Explorer)e.Control.Context;
            Outlook.Selection selected = explorer.Selection;
            string path = null;
            foreach (Object obj in selected) {
                // Outlook.MailItem mailItem = (Outlook.MailItem)obj;
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
                                        // Prompt for Outlook folder path
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
    }
}
