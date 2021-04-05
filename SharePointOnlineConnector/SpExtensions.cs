using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    internal static class SPExtensions
    {
        public static T GetFieldValue<T>(this ListItem listItem, string field, T defaultValue = default)
        {
            field = field.Replace(" ", "_x0020_");
            if (!listItem.FieldValues.ContainsKey(field))
                return defaultValue;
            var fieldValue = listItem.FieldValues[field];
            if (typeof(T) == typeof(bool))
                return (T)(object)ConvertYesNoToBool(listItem, field);
            return (T)fieldValue;
        }

        public static object GetFieldValue(this ListItem listItem, Type type, string field)
        {
            field = field.Replace(" ", "_x0020_");
            if (!listItem.FieldValues.ContainsKey(field))
                return default;
            var fieldValue = listItem.FieldValues[field];
            if (type == typeof(bool))
                return ConvertYesNoToBool(listItem, field);
            return fieldValue;
        }

        public static bool ConvertYesNoToBool(this ListItem listItem, string field)
            => listItem.GetFieldValue(field, "no").ToLower() == "yes" ? true : false;
    }
}
