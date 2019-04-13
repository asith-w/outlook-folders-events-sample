using System;
using System.Windows.Forms;
using AddinExpress.MSO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookItemsEventsDemo
{
    /// <summary>
    /// Add-in Express Outlook Folders Events Class
    /// </summary>
    public class OutlookFoldersEventsClass : AddinExpress.MSO.ADXOutlookFoldersEvents
    {
        public OutlookFoldersEventsClass(AddinExpress.MSO.ADXAddinModule module)
            : base(module)
        {
        }

        public override void ProcessFolderAdd(object folder)
        {
            if (folder != null)
            {
                Outlook.MAPIFolder newFolder = (Outlook.MAPIFolder)folder;
                Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)this.FolderObj;

                MessageBox.Show(String.Format("You've added a sub folder called {0} to {1}",
                                newFolder.Name, parentFolder.FolderPath));
            }
        }

        public override void ProcessFolderChange(object folder)
        {
            if (folder != null)
            {
                Outlook.MAPIFolder changedFolder = (Outlook.MAPIFolder)folder;
                Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)this.FolderObj;

                MessageBox.Show(String.Format("The {0} folder in {1} has changed",
                    changedFolder.Name, parentFolder.FolderPath));

                Marshal.ReleaseComObject(changedFolder);
            }
        }

        public override void ProcessFolderRemove()
        {
            Outlook.MAPIFolder parentFolder = (Outlook.MAPIFolder)this.FolderObj;
            MessageBox.Show("A subfolder has been removed from " + parentFolder.FolderPath);
        }
    }
}

