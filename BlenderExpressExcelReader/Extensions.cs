using System.Collections.Generic;

// ReSharper disable once CheckNamespace
namespace System
{
    public static class Extensions
    {
        public static T To<T>(this object value)
        {
            var type = typeof(T);

            if (!type.IsNullable()) return (T) Convert.ChangeType(value, type);
            if (value == null || value.ToString().IsNullOrWhiteSpace())
                return default(T);

            try
            {
                return (T)Convert.ChangeType(value, type.GetUnderlyingType());
            }
            catch
            {
                return (T)Convert.ChangeType(value, type.GetGenericArguments()[0]);
            }
        }
        public static bool IsNullOrWhiteSpace(this string s)
        {
            return string.IsNullOrWhiteSpace(s);
        }
        public static bool IsNullOrEmpty(this object s)
        {
            return string.IsNullOrWhiteSpace(s?.ToString());
        }
        public static bool IsNullable(this Type type)
        {
            if (type == null) throw new ArgumentNullException("type");
            if (type.IsValueType && !type.IsGenericType) return false;

            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        public static Type GetUnderlyingType(this Type type)
        {
            if (type == null) throw new ArgumentNullException("type");
            return Nullable.GetUnderlyingType(type);
        }

        public static List<T> GetList<T>()
        {
            var genericListType = typeof(List<>).MakeGenericType(typeof(T));
            return (List<T>)Activator.CreateInstance(genericListType);
        }
    }
}
