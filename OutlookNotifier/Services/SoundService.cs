using NAudio.Wave;
using System;
using System.IO;

namespace OutlookNotifier.Services
{
    public class SoundService : IDisposable
    {
        private readonly object locker = new object();

        private WaveOutEvent outputDevice;

        private AudioFileReader audioFile;

        private readonly string soundFile;

        public SoundService()
        {
            soundFile = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Sounds",
                "correo.wav");
        }

        public void Play()
        {
            lock (locker)
            {
                try
                {
                    if (!File.Exists(soundFile))
                    {
                        Logger.Write("No existe el archivo: " + soundFile);
                        return;
                    }

                    Stop();

                    audioFile = new AudioFileReader(soundFile);

                    outputDevice = new WaveOutEvent();

                    outputDevice.PlaybackStopped += OutputDevice_PlaybackStopped;

                    outputDevice.Init(audioFile);

                    outputDevice.Play();
                }
                catch (Exception ex)
                {
                    Logger.Write(ex.ToString());

                    Stop();
                }
            }
        }

        private void OutputDevice_PlaybackStopped(
            object sender,
            StoppedEventArgs e)
        {
            Stop();
        }

        public void Stop()
        {
            lock (locker)
            {
                if (outputDevice != null)
                {
                    outputDevice.PlaybackStopped -= OutputDevice_PlaybackStopped;

                    outputDevice.Dispose();

                    outputDevice = null;
                }

                if (audioFile != null)
                {
                    audioFile.Dispose();

                    audioFile = null;
                }
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }
}