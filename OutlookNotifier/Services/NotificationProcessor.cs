using Microsoft.Office.Interop.Outlook;
using OutlookNotifier.Models;
using System;

namespace OutlookNotifier.Services
{
    public class NotificationProcessor
    {
        /// <summary>
        /// Tiempo máximo (en segundos) para considerar un correo como recién recibido.
        /// Si el correo fue recibido hace más tiempo, se asume que fue movido
        /// manualmente y no se reproduce el sonido.
        /// </summary>
        private const int MaxArrivalAgeSeconds = 60;

        private readonly SoundService soundService;

        public NotificationProcessor(Application application)
        {
            soundService = new SoundService();
        }

        public void ProcessItemAdd(MailReceivedEventArgs e)
        {
            if (e == null)
                return;

            if (e.MailItem == null)
                return;

            DateTime receivedTime;

            try
            {
                receivedTime = e.MailItem.ReceivedTime;
            }
            catch
            {
                Logger.Write("[Ignorado] No fue posible obtener ReceivedTime.");
                return;
            }

            if (receivedTime == DateTime.MinValue)
            {
                Logger.Write("[Ignorado] ReceivedTime inválido.");
                return;
            }

            TimeSpan age = DateTime.Now - receivedTime;

            if (age.TotalSeconds > MaxArrivalAgeSeconds)
            {
                Logger.Write(
                    string.Format(
                        "[Ignorado] Movimiento manual | Cuenta={0} | Carpeta={1} | Asunto={2}",
                        e.Store.DisplayName,
                        e.Folder.Name,
                        e.MailItem.Subject));

                return;
            }

            Logger.Write(
                string.Format(
                    "[Correo] Cuenta={0} | Carpeta={1} | Asunto={2}",
                    e.Store.DisplayName,
                    e.Folder.Name,
                    e.MailItem.Subject));

            soundService.Play();
        }
    }
}