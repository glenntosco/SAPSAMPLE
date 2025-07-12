using System;
using System.IO;
using Newtonsoft.Json;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class App<T> where T : AppSettings, new()
    {
        private static readonly Lazy<T> _instance = new Lazy<T>(() => Activator.CreateInstance<T>().Initialize<T>());
        public static readonly T Instance = _instance.Value;
    }

    public class AppSettings
    {
        protected virtual string SettingsFileName { get; } = "settings.json";

        public T Initialize<T>() where T : AppSettings, new()
        {
            return Utils.DeserializeFromJson<T>(Utils.ReadTextFile(Path.Combine(Directory.GetCurrentDirectory(), SettingsFileName), false), null, false, new T());
        }

        public string Serialize()
        {
            return Utils.SerializeToStringJson(this, Formatting.Indented);
        }

        public void SaveToFile()
        {
            Utils.WriteTextFile(Path.Combine(Directory.GetCurrentDirectory(), SettingsFileName), Serialize());
        }
    }
}