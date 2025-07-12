using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class WorkerThread
    {
        public readonly string Name;
        private readonly Thread _thread;
        private readonly AutoResetEvent _autoReset = new AutoResetEvent(false);
        private readonly List<ScheduleSetting> _queue = new List<ScheduleSetting>();
        private static readonly object lockObj = new object();

        public WorkerThread(string name)
        {
            Name = name;
            _thread = new Thread(Run);
            _thread.Start();
        }

        public void ScheduleExecute(ScheduleSetting poll)
        {
            if (poll == null)
                return;
            lock (lockObj)
            {
                if (_queue.Any(c => c.Name == poll.Name))
                    return;
                _queue.Add(poll);
            }
            _autoReset.Set();
        }

        private bool _pendingStop;
        public void Stop()
        {
            lock (lockObj)
                _queue.Clear();

            _pendingStop = true;
            _autoReset.Set();
            _thread.Join(TimeSpan.FromSeconds(10));
            if (!_thread.IsAlive)
                return;

            _thread.Abort("Stop grace period expired");
            _thread.Join(TimeSpan.FromSeconds(10));
        }

        private void Run()
        {
            ScheduleThread.Report($"Thread: [{Name}] started");
            while (!_pendingStop)
            {
                try
                {
                    bool sleep;
                    lock (lockObj)
                        sleep = !_queue.Any();
                    if (sleep)
                    {
                        _autoReset.WaitOne(TimeSpan.FromSeconds(10));
                        continue;
                    }

                    ScheduleSetting task;
                    lock (lockObj)
                    {
                        task = _queue[0];
                        _queue.RemoveAt(0);
                    }

                    lock (lockObj)
                        Console.Out.WriteLine($"{_queue.Count} Thread: [{Name}]: Executing {task.Name}");

                    var type = Assembly.GetEntryAssembly()?.GetType(task.Class);
                    if (type == null)
                        throw new Exception($"Thread: [{Name}]: Invalid type: [{task.Class}]");
                    if (!(Activator.CreateInstance(type, task) is BaseWorker worker))
                        return;
                    try
                    {
                        worker.Execute();
                    }
                    catch (Exception e)
                    {
                        worker.LogError(new Exception($"Thread: [{Name}]", e));
                    }
                }
                catch (Exception e)
                {
                    try { ScheduleThread.ReportError(e); }
                    catch (Exception ex) { Console.WriteLine(ex); }
                }
            }
            ScheduleThread.Report($"Thread: [{Name}] stopped");
        }
    }
}