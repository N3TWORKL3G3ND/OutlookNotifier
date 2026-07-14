using NAudio.CoreAudioApi;
using System;
using System.Diagnostics;

namespace OutlookNotifier.Services
{
    public class VolumeService
    {
        private readonly MMDevice device;

        public VolumeService()
        {
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();

            device = enumerator.GetDefaultAudioEndpoint(
                DataFlow.Render,
                Role.Multimedia);
        }

        public float GetOutlookVolume()
        {
            SessionCollection sessions = device.AudioSessionManager.Sessions;

            for (int i = 0; i < sessions.Count; i++)
            {
                AudioSessionControl session = sessions[i];

                try
                {
                    int pid = (int)session.GetProcessID;

                    Process process = Process.GetProcessById(pid);

                    if (process.ProcessName.Equals(
                        "OUTLOOK",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return session.SimpleAudioVolume.Volume;
                    }
                }
                catch
                {
                    // Ignorar sesiones inválidas
                }
            }

            return 1.0f;
        }

        public float SetOutlookVolume(float volume)
        {
            SessionCollection sessions = device.AudioSessionManager.Sessions;

            for (int i = 0; i < sessions.Count; i++)
            {
                AudioSessionControl session = sessions[i];

                try
                {
                    int pid = (int)session.GetProcessID;

                    Process process = Process.GetProcessById(pid);

                    if (process.ProcessName.Equals(
                        "OUTLOOK",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        SimpleAudioVolume audio = session.SimpleAudioVolume;

                        float original = audio.Volume;

                        audio.Volume = volume;

                        return original;
                    }
                }
                catch
                {
                    // Ignorar sesiones inválidas
                }
            }

            return 1.0f;
        }
    }
}