using Microsoft.Office.Interop.Outlook;
using OutlookNotifier.Models;
using System.Collections.Generic;

namespace OutlookNotifier.Services
{
    public class NotificationEngine
    {
        private readonly Application application;

        private readonly List<FolderMonitor> monitors =
            new List<FolderMonitor>();

        private readonly NotificationProcessor processor;

        private readonly HashSet<string> ignoredFolders =
            new HashSet<string>();

        public NotificationEngine(Application application)
        {
            this.application = application;

            processor = new NotificationProcessor(application);

            InitializeIgnoredFolders();
        }

        public void Start()
        {
            NameSpace session = application.Session;

            foreach (Store store in session.Stores)
            {
                ScanFolder(store.GetRootFolder());
            }

            Logger.Write("NotificationEngine iniciado.");
        }

        private void InitializeIgnoredFolders()
        {
            NameSpace session = application.Session;

            ignoredFolders.Add(
                session.GetDefaultFolder(
                    OlDefaultFolders.olFolderDrafts).EntryID);

            ignoredFolders.Add(
                session.GetDefaultFolder(
                    OlDefaultFolders.olFolderDeletedItems).EntryID);

            ignoredFolders.Add(
                session.GetDefaultFolder(
                    OlDefaultFolders.olFolderOutbox).EntryID);

            ignoredFolders.Add(
                session.GetDefaultFolder(
                    OlDefaultFolders.olFolderSentMail).EntryID);
        }

        private void ScanFolder(MAPIFolder folder)
        {
            bool monitorFolder =
                folder.DefaultItemType == OlItemType.olMailItem &&
                !ignoredFolders.Contains(folder.EntryID);

            if (monitorFolder)
            {
                FolderMonitor monitor = new FolderMonitor(folder);

                monitor.MailReceived += Monitor_MailReceived;

                monitors.Add(monitor);
            }

            foreach (MAPIFolder child in folder.Folders)
            {
                ScanFolder(child);
            }
        }

        private void Monitor_MailReceived(
            object sender,
            MailReceivedEventArgs e)
        {
            processor.ProcessItemAdd(e);
        }
    }
}