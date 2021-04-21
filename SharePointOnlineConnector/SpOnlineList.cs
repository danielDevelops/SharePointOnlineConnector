using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public class SpOnlineList<T> where T : class, new()
    {
        private readonly SharePointContext connector;
        private readonly string listName;
        private readonly IEnumerable<PropertyAttribute> attributesOfT;

        public SpOnlineList(SharePointContext connector)
        {
            this.connector = connector;
            var attribute = typeof(T).GetCustomAttributes(typeof(SpListAttribute), false)?.Cast<SpListAttribute>()?.SingleOrDefault();
            this.listName = $"{typeof(T).Name}s";
            if (!string.IsNullOrWhiteSpace(attribute?.ListName))
                this.listName = attribute.ListName;
            attributesOfT = PropertyAttribute.GetPropertyAttributesForT(typeof(T));
        }

        public SpOnlineList(SharePointContext connector, string listName)
        {
            this.connector = connector;
            this.listName = listName;
            attributesOfT = PropertyAttribute.GetPropertyAttributesForT(typeof(T));
        }

        public async Task<IEnumerable<T>> GetItemsAsync(string camlQuery, Func<T, Dictionary<string, object>, Task> doAfterLookup = null)
        {
            var web = await connector.InitWebAsync();
            var list = web.Lists.GetByTitle(listName);
            var caml = new CamlQuery();
            caml.ViewXml = camlQuery;
            var colListItems = list.GetItems(caml);
            connector.Context.Load(colListItems);
            await connector.Context.ExecuteQueryAsync();
            var itemList = new List<T>();
            foreach (var item in colListItems)
            {
                var itemT = new T();
                foreach (var prop in attributesOfT)
                {
                    if (prop.IsEmail)
                        prop.Property.SetValue(itemT, item.GetFieldValue<FieldUserValue>(prop.ColumnName)?.Email);
                    else
                        prop.Property.SetValue(itemT, item.GetFieldValue(prop.FieldType, prop.ColumnName));
                }

                if (doAfterLookup != null)
                {
                    var itemsToUpdate = new Dictionary<string, object>();
                    await doAfterLookup(itemT, itemsToUpdate);
                    foreach (var update in itemsToUpdate)
                    {
                        item[update.Key] = update.Value;
                    }
                    if (itemsToUpdate.Any())
                    {
                        item.Update();
                        await connector.Context.ExecuteQueryAsync();
                    }
                }
                itemList.Add(itemT);
            }
            return itemList;
        }

        public async Task InsertItemAsync(T value)
        {
            var web = await connector.InitWebAsync();
            var list = web.Lists.GetByTitle(listName);
            var listItemCreateInfo = new ListItemCreationInformation();
            var newItem = list.AddItem(listItemCreateInfo);

            foreach (var prop in attributesOfT)
            {
                var propVal = prop.Property.GetValue(value);
                if (propVal == null ||
                  prop.FieldName == "Author" ||
                  prop.FieldName == "Editor" ||
                  prop.FieldName == "Created" ||
                  prop.FieldName == "Modified")
                    continue;
                newItem[prop.ColumnName] = propVal;
            }

            newItem.Update();
            await connector.Context.ExecuteQueryAsync();
        }
    }
}
