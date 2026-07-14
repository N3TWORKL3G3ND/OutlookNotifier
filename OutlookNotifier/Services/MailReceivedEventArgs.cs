using Microsoft.Office.Interop.Outlook;
using System;

namespace OutlookNotifier.Models
{
    public class MailReceivedEventArgs : EventArgs
    {
        public MailItem MailItem { get; }

        public MAPIFolder Folder { get; }

        public Store Store { get; }

        public MailReceivedEventArgs(
            MailItem mailItem,
            MAPIFolder folder,
            Store store)
        {
            MailItem = mailItem;
            Folder = folder;
            Store = store;
        }
    }
}