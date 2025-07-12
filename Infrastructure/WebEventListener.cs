using System;
using System.Collections.Concurrent;
using System.Net;
using System.Text;
using Microsoft.AspNet.SignalR.Client;
using Microsoft.AspNet.SignalR.Client.Hubs;
using RestSharp;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class WebEventListener
    {
        private static readonly object _syncObject = new object();

        private HubConnection _hub;
        private IHubProxy _proxy;

        private readonly ConcurrentDictionary<string, Action<dynamic>> _subscribers = new ConcurrentDictionary<string, Action<dynamic>>();

        public void Subscribe(string eventName, Action<dynamic> callback)
        {
            _subscribers[eventName] = callback;
            _proxy?.On(eventName, payload =>
            {
                _subscribers[eventName]?.Invoke(payload);
            });
        }

        public void Start()
        {
            lock (_syncObject)
            {
                _hub = _hub ?? new HubConnection(Singleton<EntryPoint>.Instance.CloudUrl, true)
                {
                    TraceLevel = TraceLevels.StateChanges,
                    TraceWriter = Singleton<LogTraceWriter>.Instance
                };

                if (_proxy == null)
                {
                    _proxy = _hub.CreateHubProxy("SubscriberHub") as HubProxy;
                    foreach (var key in _subscribers.Keys)
                        _proxy.On(key, payload => { _subscribers[key]?.Invoke(payload); });
                }

                _hub.Closed += () =>
                {
                    _hub = null;
                    _proxy = null;
                    Start();
                };

                _hub.StateChanged += s =>
                {
                    if (s.NewState == ConnectionState.Connected)
                        _proxy.Invoke("JoinGroup", Singleton<EntryPoint>.Instance.DeploymentId);
                };

                if (_hub.State == ConnectionState.Disconnected)
                {
                    try
                    {
                        _hub.Start();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
        }

        public void Stop()
        {
            _hub?.Stop(TimeSpan.FromSeconds(5));
        }
    }

    public class LogTraceWriter : System.IO.TextWriter
    {
        public override void WriteLine(string value)
        {
            if (Singleton<EntryPoint>.Instance.DeploymentId != null)
                Singleton<EntryPoint>.Instance.WebInvoke("api/WorkerDeploymentApi/WriteToConsole", null, Method.POST, new
                {
                    WorkerDeploymentId = Singleton<EntryPoint>.Instance.DeploymentId,
                    Pid = ScheduleThread.Process.Value.Id,
                    WorkerName = "SignalR",
                    ScheduleThread.Process.Value.ProcessName,
                    MachineName = Dns.GetHostName(),
                    Timestamp = DateTimeOffset.Now,
                    Data = value,
                    StdType = "StdOut"
                }, true);
            Console.Out.WriteLine(value);
        }

        public override Encoding Encoding => Encoding.Default;
    }
}