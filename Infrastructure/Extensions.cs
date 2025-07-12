using System;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public static class Extensions
    {
        public static string Truncate(this string value, int maxLength)
        {
            return string.IsNullOrEmpty(value) ? value : value.Substring(0, Math.Min(value.Length, maxLength));
        }

        public static T ParseEnum<T>(this string value, bool throwExc = true, T defaultVal = default)
        {
            try
            {
                if (value == null && Nullable.GetUnderlyingType(typeof(T)) != null)
                    return defaultVal;
                return (T)Enum.Parse(typeof(T), value ?? string.Empty, true);
            }
            catch
            {
                if (throwExc)
                    throw;
                return defaultVal;
            }
        }
    }
}