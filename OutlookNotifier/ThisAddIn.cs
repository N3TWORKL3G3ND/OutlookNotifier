using OutlookNotifier.Services;
using System;

namespace OutlookNotifier
{
    public partial class ThisAddIn
    {
        private NotificationEngine notificationEngine;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            notificationEngine = new NotificationEngine(Application);

            notificationEngine.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region Código generado por VSTO

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}