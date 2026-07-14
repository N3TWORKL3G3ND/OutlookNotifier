using System;
using System.IO;

namespace OutlookNotifier.Services
{
    public static class Logger
    {
        private static readonly object locker = new object();

        private static readonly string logPath =
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments),
                "OutlookNotifier.log");

        public static void Write(string message)
        {
            try
            {
                lock (locker)
                {
                    File.AppendAllText(
                        logPath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") +
                        "  " +
                        message +
                        Environment.NewLine);
                }
            }
            catch
            {
                // Nunca permitir que un error del log afecte Outlook.
            }
        }
    }
}