using Microsoft.Office.Interop.Outlook;
using OutlookNotifier.Models;
using System;

namespace OutlookNotifier.Services
{
    public class FolderMonitor
    {
        private readonly MAPIFolder folder;

        private readonly Items items;

        public event EventHandler<MailReceivedEventArgs> MailReceived;

        public FolderMonitor(MAPIFolder folder)
        {
            this.folder = folder;

            items = folder.Items;

            items.ItemAdd += Items_ItemAdd;
        }

        private void Items_ItemAdd(object item)
        {
            MailItem mail = item as MailItem;

            if (mail == null)
                return;

            MailReceived?.Invoke(
                this,
                new MailReceivedEventArgs(
                    mail,
                    folder,
                    folder.Store));
        }
    }
}