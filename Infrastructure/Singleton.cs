using System;
using System.Linq;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class Singleton<T> where T : class, new()
    {
        private static readonly Lazy<T> _instance = new Lazy<T>(() => Factory.Create<T>());
        public static readonly T Instance = _instance.Value;
    }

    public class Factory
    {
        public static T Create<T>(params object[] pars) where T : class
        {
            return typeof(T).GetConstructors().SingleOrDefault(c => c.GetParameters().Length == pars.Length)?.Invoke(pars) as T;
        }
    }
}