using System;
using System.Diagnostics;
using System.Net;
using RestSharp;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public abstract class BaseWorker
    {
        private static Lazy<Process> Process => new Lazy<Process>(System.Diagnostics.Process.GetCurrentProcess);
        protected readonly ScheduleSetting _settings;

        protected BaseWorker(ScheduleSetting settings)
        {
            _settings = settings;
        }

        public abstract void Execute();

        public void Log(string msg)
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
                Singleton<EntryPoint>.Instance.WebInvoke("api/WorkerDeploymentApi/WriteToConsole", null, Method.POST, new
                {
                    WorkerDeploymentId = Singleton<EntryPoint>.Instance.DeploymentId,
                    Pid = Process.Value.Id,
                    WorkerName = GetType().Name,
                    Process.Value.ProcessName,
                    MachineName = Dns.GetHostName(),
                    Timestamp = DateTimeOffset.Now,
                    Data = msg,
                    StdType = "StdOut"
                });
            Console.Out.WriteLine(msg);
        }

        public void LogError(Exception ex)
        {
            LogError(ex.ToString());
        }

        public void LogError(string msg)
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
                Singleton<EntryPoint>.Instance.WebInvoke("api/WorkerDeploymentApi/WriteToConsole", null, Method.POST, new
                {
                    WorkerDeploymentId = Singleton<EntryPoint>.Instance.DeploymentId,
                    Pid = Process.Value.Id,
                    WorkerName = GetType().Name,
                    Process.Value.ProcessName,
                    MachineName = Dns.GetHostName(),
                    Timestamp = DateTimeOffset.Now,
                    Data = msg,
                    StdType = "StdError"
                });
            Console.Error.WriteLine(msg);
        }
    }
}