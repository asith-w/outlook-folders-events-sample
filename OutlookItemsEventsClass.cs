using System;
using System.Windows.Forms;
using AddinExpress.MSO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutlookItemsEventsDemo
{
    /// <summary>
    /// Add-in Express Outlook Items Events Class
    /// </summary>
    public class OutlookItemsEventsClass : AddinExpress.MSO.ADXOutlookItemsEvents
    {
        public OutlookItemsEventsClass(AddinExpress.MSO.ADXAddinModule module)
            : base(module)
        {
        }

        public override void ProcessItemAdd(object item)
        {
            if (item is Outlook._ContactItem)
            {
                Outlook._ContactItem newContact = (Outlook._ContactItem)item;
                Outlook._Application outlookApp = ((AddinModule)this.Module).OutlookApp;
                Outlook.MAPIFolder targetFolder = (Outlook.MAPIFolder)this.FolderObj;
                Outlook._JournalItem journalItem = null;

                try
                {
                    journalItem = (Outlook._JournalItem)outlookApp.CreateItem(Outlook.OlItemType.olJournalItem);
                    journalItem.Subject = String.Format("You've added {0} to {1}",
                        newContact.FullName, targetFolder.FolderPath);
                    journalItem.Save();
                }
                finally
                {
                    if (journalItem != null)
                        Marshal.ReleaseComObject(journalItem);
                }
            }
        }

        public override void ProcessItemChange(object item)
        {
            if (item is Outlook._ContactItem)
            {
                Outlook._ContactItem changedContact = (Outlook._ContactItem)item;

                if (String.IsNullOrEmpty(changedContact.JobTitle))
                {
                    MessageBox.Show(String.Format("Please add a job title for {0}. All contacts must have proper titles",
                                    changedContact.FullName), "Title Required");
                }
            }
        }

        public override void ProcessItemRemove()
        {
            Outlook._Application outlookApp = ((AddinModule)this.Module).OutlookApp;
            Outlook.MAPIFolder targetFolder = (Outlook.MAPIFolder)this.FolderObj;
            Outlook._JournalItem journalItem = null;

            try
            {
                journalItem = (Outlook._JournalItem)outlookApp.CreateItem(Outlook.OlItemType.olJournalItem);
                journalItem.Subject = String.Format("You've removed items from {0}", targetFolder.FolderPath);
                journalItem.Save();
            }
            finally
            {
                if (journalItem != null)
                    Marshal.ReleaseComObject(journalItem);
            }
        }

        public override void ProcessBeforeItemMove(object item, object moveTo, AddinExpress.MSO.ADXCancelEventArgs e)
        {
            if (item is Outlook._ContactItem)
            {
                Outlook._ContactItem movedContact = (Outlook._ContactItem)item;
                Outlook._Application outlookApp = ((AddinModule)this.Module).OutlookApp;
                Outlook._JournalItem journalItem = null;

                try
                {
                    journalItem = (Outlook._JournalItem)outlookApp.CreateItem(Outlook.OlItemType.olJournalItem);
                    if (moveTo != null)
                    {
                        Outlook.MAPIFolder targetFolder = (Outlook.MAPIFolder)moveTo;
                        journalItem.Subject = String.Format("You've moved {0} to {1}",
                            movedContact.FullName, targetFolder.FolderPath);
                        Marshal.ReleaseComObject(moveTo);
                    }
                    else
                    {
                        journalItem.Subject = String.Format("You've permanently deleted {0}",
                            movedContact.FullName);
                    }
                    journalItem.Save();
                }
                finally
                {
                    if (journalItem != null)
                        Marshal.ReleaseComObject(journalItem);
                }
            }
        }

        public override void ProcessBeforeFolderMove(object moveTo, AddinExpress.MSO.ADXCancelEventArgs e)
        {
            if (this.FolderObj != null)
            {
                Outlook._Application outlookApp = ((AddinModule)this.Module).OutlookApp;
                Outlook._JournalItem journalItem = null;

                try
                {
                    journalItem = (Outlook._JournalItem)outlookApp.CreateItem(Outlook.OlItemType.olJournalItem);
                    if (moveTo != null)
                    {
                        Outlook.MAPIFolder targetFolder = (Outlook.MAPIFolder)moveTo;
                        journalItem.Subject = String.Format("You've moved the folder to {1}",
                                targetFolder.FolderPath);
                    }
                    else
                    {
                        Outlook.MAPIFolder deletedFolder = (Outlook.MAPIFolder)this.FolderObj;
                        journalItem.Subject = String.Format("You've permanently deleted {0}",
                            deletedFolder.FolderPath);
                    }
                }
                finally
                {
                    if (journalItem != null)
                        Marshal.ReleaseComObject(journalItem);
                }
            }
        }

    }
}

