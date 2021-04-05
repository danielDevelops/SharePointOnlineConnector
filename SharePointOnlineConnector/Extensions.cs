using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public static class Extensions
    {
        public static SpOnlineList<T> InitSpOnlineListOfT<T>(this Connector connector, string listName) where T : class, new()
            => new SpOnlineList<T>(connector, listName);
        public static string WrapCaml(this string caml, string wrapper)
            => $"<{wrapper}>{caml}</{wrapper}>";
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ItemAttribute : Attribute
    {
        public string FieldName;
        public bool IsEmail;
        public bool Ignore;
        public ItemAttribute(string fieldName)
        {
            FieldName = fieldName;
            IsEmail = false;
            Ignore = false;
        }
    }
}
