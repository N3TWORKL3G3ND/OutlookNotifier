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
        }

        public void Start()
        {
            NameSpace session = application.Session;

            foreach (Store store in session.Stores)
            {
                InitializeIgnoredFolders(store);

                ScanFolder(store.GetRootFolder());
            }

            Logger.Write("NotificationEngine iniciado.");
        }

        private void InitializeIgnoredFolders(Store store)
        {
            AddIgnoredFolder(store, OlDefaultFolders.olFolderDrafts);
            AddIgnoredFolder(store, OlDefaultFolders.olFolderDeletedItems);
            AddIgnoredFolder(store, OlDefaultFolders.olFolderOutbox);
            AddIgnoredFolder(store, OlDefaultFolders.olFolderSentMail);
        }

        private void AddIgnoredFolder(
            Store store,
            OlDefaultFolders folderType)
        {
            try
            {
                MAPIFolder folder = store.GetDefaultFolder(folderType);

                if (folder != null)
                {
                    ignoredFolders.Add(folder.EntryID);
                }
            }
            catch
            {
                // Algunos Stores (PST, IMAP, buzones compartidos, etc.)
                // pueden no contener todas las carpetas predeterminadas.
            }
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