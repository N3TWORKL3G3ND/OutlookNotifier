using System;
using System.Collections.Generic;

namespace OutlookNotifier.Services
{
    public class DuplicateDetector
    {
        private readonly Dictionary<string, DateTime> cache =
            new Dictionary<string, DateTime>();

        private readonly object sync = new object();

        public bool IsDuplicate(string id)
        {
            lock (sync)
            {
                DateTime now = DateTime.Now;

                List<string> remove = new List<string>();

                foreach (KeyValuePair<string, DateTime> item in cache)
                {
                    if ((now - item.Value).TotalSeconds > 10)
                        remove.Add(item.Key);
                }

                foreach (string key in remove)
                    cache.Remove(key);

                if (cache.ContainsKey(id))
                    return true;

                cache.Add(id, now);

                return false;
            }
        }
    }
}