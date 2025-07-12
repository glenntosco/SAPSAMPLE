using System;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using RestSharp;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class EntryPoint
    {
        public Guid? DeploymentId { get; private set; }
        public string ApiKey { get; set; }
        public string CloudUrl { get; set; }

        private RestClient _client;
        public RestClient Client => _client = _client ?? new RestClient(CloudUrl);

        public void Startup(string[] args)
        {
            DeploymentId = Guid.TryParse(args.Length > 0 ? args[0] : ConfigurationManager.AppSettings["DeploymentId"], out var depId) ? depId : (Guid?)null;
            ApiKey = args.Length > 1 ? args[1] : ConfigurationManager.AppSettings["ApiKey"];
            CloudUrl = args.Length > 2 ? args[2] : ConfigurationManager.AppSettings["CloudUrl"];

            Console.Out.WriteLine($"CloudUrl: {CloudUrl}");
            Console.Out.WriteLine($"DeploymentId: {DeploymentId}");
            Console.Out.WriteLine($"ApiKey: {ApiKey}");
            Start();
            Console.In.ReadLine();
            Stop();
        }

        protected virtual void Start()
        {
            ScheduleThread.Instance.Start();
        }

        protected virtual void Stop()
        {
            ScheduleThread.Instance.Stop();
        }

        public IRestResponse WebInvoke(string url, Action<RestRequest> requestRewrite = null, Method method = Method.GET, object payload = null, bool async = false)
        {
            try
            {
                var request = new RestRequest(url, method) { RequestFormat = DataFormat.Json };
                requestRewrite?.Invoke(request);
                request.AddHeader("ApiKey", ApiKey);
                if (method == Method.POST || method == Method.PATCH)
                    request.AddJsonBody(payload);
                if (async)
                {
                    Client.ExecuteAsync(request);
                    return null;
                }
                var resp = Client.Execute(request);

                if (!resp.IsSuccessful)
                    throw new BusinessWebException(resp.StatusCode, resp.Content);
                return resp;
            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.ToString());
                return null;
            }
        }

        public async Task WebInvokeAsync(string url, Method method = Method.GET, object payload = null)
        {
            var result = await WebInvokeResponseAsync(url, method, payload);
            if (!result.IsSuccessful)
                throw new BusinessWebException(result.StatusCode, result.Content);
        }

        public async Task<T> WebInvokeAsync<T>(string url, string root = null, Method method = Method.GET, object payload = null) where T : class
        {
            var result = await WebInvokeResponseAsync(url, method, payload);
            if (result.IsSuccessful)
                return Utils.DeserializeFromJson<T>(result.Content, root);

            throw new BusinessWebException(result.StatusCode, result.Content);
        }

        private Task<IRestResponse> WebInvokeResponseAsync(string url, Method method = Method.GET, object payload = null)
        {
            var split = url.Replace("\r", string.Empty).Replace("\n", string.Empty).Split('?');
            if (split.Length > 1)
                url = split[0];
            var request = new RestRequest(url, method) { RequestFormat = DataFormat.Json };
            request.AddHeader("ApiKey", ApiKey);
            if (split.Length > 1)
            {
                var qryParsString = split.Last();
                var parsed = HttpUtility.ParseQueryString(qryParsString);
                foreach (var par in parsed.Cast<string>().Select(c => c.Trim()).Where(c => !string.IsNullOrWhiteSpace(c)))
                    request.AddQueryParameter(par, parsed.Get(par));
            }

            if (method == Method.POST || method == Method.PATCH)
                request.AddJsonBody(payload);

            return Client.ExecuteAsync(request);
        }
    }
}