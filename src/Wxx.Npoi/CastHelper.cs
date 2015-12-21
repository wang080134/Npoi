using System;

namespace Wxx.Npoi
{
    public static class CastHelper
    {
        public static object CastTo(object value, Type conversionType)
        {
            if (value == null)
                return value;
            if (IsNullableType(conversionType))
                conversionType = Nullable.GetUnderlyingType(conversionType);
            if (conversionType.IsEnum)
                return Enum.Parse(conversionType, value.ToString(), true);
            if (conversionType == typeof(Guid))
                return Guid.Parse(value.ToString());
            if (conversionType == typeof(bool))
            {
                switch (value.ToString().ToLower())
                {
                    case "0":
                        return false;
                    case "1":
                        return true;
                    case "是":
                        return true;
                    case "否":
                        return false;
                    case "yes":
                        return true;
                    case "no":
                        return false;
                    case "true":
                        return true;
                    case "false":
                        return false;
                }
            }
            if (value is IConvertible)
                return Convert.ChangeType(value, conversionType);
            return value;
        }

        /// <summary>
        /// 把对象实例转化为指定类型
        /// </summary>
        /// <typeparam name="T">动态类型</typeparam>
        /// <param name="value">要转化的源对象</param>
        /// <returns>转化后的指定类型的对象，转化失败引发异常。</returns>
        public static T CastTo<T>(object value)
        {
            object result = CastTo(value, typeof(T));
            return (T)result;
        }

        public static T CastTo<T>(object value, T defaultValue)
        {
            try
            {
                T res = CastTo<T>(value);
                Type type = typeof(T);
                if (IsNullableType(type) && res == null)
                    return defaultValue;
                if (!type.IsValueType && res == null)
                    return defaultValue;
                return res;
            }
            catch
            {
                return defaultValue;
            }
        }
        
        public static bool IsNullableType(Type type)
        {
            bool result = false;
            if (type != null &&
                type.IsGenericType &&  // 是否为泛型
                type.GetGenericTypeDefinition() == typeof(Nullable<>)  // 是否为 Nullable<> 的定义类型
                )
            {
                result = true;
            }
            return result;
        }
    }
}
