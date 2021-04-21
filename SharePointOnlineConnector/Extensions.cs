using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public static class Extensions
    {
        internal static bool DynamicHasProperty(dynamic prop, string property)
            => ((IDictionary<string, object>)prop).ContainsKey(property);

        public static string WrapCaml(this string caml, string wrapper)
            => $"<{wrapper}>{caml}</{wrapper}>";
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class SpPropertyAttribute : Attribute
    {
        public string FieldName;
        public bool IsEmail;
        public bool Ignore;
        public SpPropertyAttribute(string fieldName)
        {
            FieldName = fieldName;
            IsEmail = false;
            Ignore = false;
        }
    }

    [AttributeUsage(AttributeTargets.Class)]
    public class SpListAttribute : Attribute
    {
        public string ListName;
        public SpListAttribute(string listName)
        {
            ListName = listName;
        }
    }
}
