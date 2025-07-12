using Pro4Soft.SapB1Integration.Infrastructure;

namespace Pro4Soft.SapB1Integration
{
    class Program
    {
        static void Main(string[] args)
        {
            Singleton<EntryPoint>.Instance.Startup(args);
        }
    }
}
