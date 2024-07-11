using System.Reflection;
using Introl.Timesheets.Api.Attributes;

namespace Introl.Timesheets.Api.Extensions;

public static class EnumExtensions
{
    public static string StringValue<T>(this T value)
        where T : Enum
    {
        var fieldName = value.ToString();
        var field = typeof(T).GetField(fieldName, BindingFlags.Public | BindingFlags.Static);
        return field?.GetCustomAttribute<StringValueAttribute>()?.Value ?? fieldName;
    }
}
