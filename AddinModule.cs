using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookItemsEventsDemo
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("CD0CE3CF-446E-4D8C-B5BC-8BC414A52B8C"), ProgId("OutlookItemsEventsDemo.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {


        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler


        }

        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;

        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            // 
            // AddinModule
            // 
            this.AddinName = "OutlookItemsEventsDemo";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
            this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);
            this.AddinBeginShutdown += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinBeginShutdown);

        }
        #endregion

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Outlook._Application OutlookApp
        {
            get
            {
                return (HostApplication as Outlook._Application);
            }
        }

        OutlookItemsEventsClass itemsEvents = null;
        OutlookFoldersEventsClass folderEvents = null;

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            itemsEvents = new OutlookItemsEventsClass(this);
            itemsEvents.ConnectTo(ADXOlDefaultFolders.olFolderContacts, true);

            folderEvents = new OutlookFoldersEventsClass(this);
            Outlook.NameSpace ns = OutlookApp.GetNamespace("MAPI");
            Outlook.MAPIFolder folderContacts = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);                        
            folderEvents.ConnectTo(folderContacts, true);
            if (ns != null) Marshal.ReleaseComObject(ns);
        }

        private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
        {
            if (itemsEvents != null)
            {
                itemsEvents.RemoveConnection();
                itemsEvents.Dispose();
            }

            if (folderEvents != null)
            {
                folderEvents.RemoveConnection();
                folderEvents.Dispose();
            }
        }

    }
}

