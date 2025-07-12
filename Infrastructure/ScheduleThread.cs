using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Threading;
using RestSharp;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class ScheduleThread
    {
        private static ScheduleThread _instance;
        public static ScheduleThread Instance => _instance = _instance ?? new ScheduleThread();

        internal static Lazy<Process> Process => new Lazy<Process>(System.Diagnostics.Process.GetCurrentProcess);

        readonly AutoResetEvent _resetEvent = new AutoResetEvent(false);

        //private WebEventListener _listener;

        private bool _isRunning = true;
        private Thread _scheduleThread;
        private List<WorkerThread> _workerThreads = new List<WorkerThread>();

        private void RunSchedulling()
        {
            _workerThreads.ForEach(c => c.Stop());
            _workerThreads = App<Settings>.Instance.Schedules.Select(c => c.ThreadName).Distinct().Select(c => new WorkerThread(c)).ToList();

            foreach (var taskOnStartup in App<Settings>.Instance.Schedules
                .Where(c => c.Active)
                .Where(c => c.RunOnStartup))
                RunTask(taskOnStartup);

            while (_isRunning)
            {
                var now = DateTimeOffset.Now;
                var (sleepTime, tasksToExecute) = App<Settings>.Instance.NextTime(now);
                if (!tasksToExecute.Any())
                {
                    Report($"Nothing to execute");
                    break;
                }

                _resetEvent.WaitOne(sleepTime);
                if (!_isRunning)
                    break;

                foreach (var poll in tasksToExecute)
                    RunTask(poll);
            }

            Report($"Stopped");
        }

        public void RunTask(ScheduleSetting task)
        {
            if (task == null)
                return;
            _workerThreads.SingleOrDefault(c => c.Name == task.ThreadName)?.ScheduleExecute(task);
        }

        internal static void Report(string msg, string workerName = null)
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
                Singleton<EntryPoint>.Instance.WebInvoke("api/WorkerDeploymentApi/WriteToConsole", null, Method.POST, new
                {
                    WorkerDeploymentId = Singleton<EntryPoint>.Instance.DeploymentId,
                    Pid = Process.Value.Id,
                    WorkerName = workerName,
                    Process.Value.ProcessName,
                    MachineName = Dns.GetHostName(),
                    Timestamp = DateTimeOffset.Now,
                    Data = msg,
                    StdType = "StdOut"
                });
            Console.Out.WriteLine(msg);
        }

        internal static void ReportError(Exception ex)
        {
            ReportError(ex.ToString());
        }

        internal static void ReportError(string msg, string workerName = null)
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
                Singleton<EntryPoint>.Instance.WebInvoke("api/WorkerDeploymentApi/WriteToConsole", null, Method.POST, new
                {
                    WorkerDeploymentId = Singleton<EntryPoint>.Instance.DeploymentId,
                    Pid = Process.Value.Id,
                    WorkerName = workerName,
                    Process.Value.ProcessName,
                    MachineName = Dns.GetHostName(),
                    Timestamp = DateTimeOffset.Now,
                    Data = msg,
                    StdType = "StdError"
                });
            Console.Error.WriteLine(msg);
        }

        public void Start()
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
            {
                Singleton<WebEventListener>.Instance.Subscribe($"ExecTask", (payload) =>
                {
                    try
                    {
                        var task = App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Name == payload);
                        if (task == null)
                            return;

                        Report($"Executing on demand", task.Class.Split('.').LastOrDefault());
                        RunTask(task);
                    }
                    catch (Exception ex)
                    {
                        ReportError(ex);
                    }
                });
                Singleton<WebEventListener>.Instance.Start();
            }

            _scheduleThread = new Thread(RunSchedulling);
            _scheduleThread.Start();
        }

        public void Stop()
        {
            _isRunning = false;
            _resetEvent.Set();
            _scheduleThread.Join();
            _workerThreads.ForEach(c => c.Stop());
            _workerThreads = new List<WorkerThread>();
        }
    }
}