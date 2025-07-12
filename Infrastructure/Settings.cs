using System;
using System.Collections.Generic;
using System.Linq;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class Settings : AppSettings
    {
        public List<ScheduleSetting> Schedules { get; set; } = new List<ScheduleSetting>();
        public (TimeSpan, List<ScheduleSetting>) NextTime(DateTimeOffset now)
        {
            var smallestTimespan = TimeSpan.MaxValue;
            var schedules = new List<ScheduleSetting>();
            foreach (var schedule in Schedules
                .Where(c => c.Start != null)
                .Where(c => c.Sleep != null)
                .Where(c => c.Active))
            {
                var pollTimeout = schedule.GetTimeout(now);
                if (smallestTimespan == pollTimeout)
                    schedules.Add(schedule);
                else if (smallestTimespan > pollTimeout)
                {
                    smallestTimespan = pollTimeout;
                    schedules = new List<ScheduleSetting> { schedule };
                }
            }
            return (smallestTimespan, schedules);
        }
    }

    public class ScheduleSetting
    {
        public string Name { get; set; }
        public bool Active { get; set; }
        public bool RunOnStartup { get; set; }
        public DateTimeOffset? Start { get; set; }
        public TimeSpan? Sleep { get; set; }
        public string Class { get; set; }
        public string ThreadName { get; set; } = "Default";

        public TimeSpan GetTimeout(DateTimeOffset now)
        {
            if (Sleep == null)
                return TimeSpan.MaxValue;

            Start = Start ?? now;

            var totalMillis = (long)now.Subtract(Start.Value).TotalMilliseconds;
            var wholeIntervals = totalMillis / (long)Sleep?.TotalMilliseconds;
            var nextRun = Start?.Add(TimeSpan.FromMilliseconds((wholeIntervals + 1) * (long)Sleep?.TotalMilliseconds));
            return nextRun.Value.Subtract(now);
        }

        public Dictionary<string, string> AdditionalSettings { get; set; } = new Dictionary<string, string>();

        public string this[string key] => AdditionalSettings.TryGetValue(key, out var res) ? res : null;
    }
}