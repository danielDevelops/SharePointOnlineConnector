using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public abstract class SharePointContext : IDisposable
    {
        private Web web;
        private readonly AuthenticationManager authManager;
        internal ClientContext Context { get; }

        public SharePointContext(string adAppId, string url, string username, string password)
        {
            authManager = new AuthenticationManager(adAppId);
            var uri = new Uri($"{url}");
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            Context = authManager.GetContext(uri, username, securePassword);
            FindAndInitializeSpOnlineLists();
        }

        internal async Task<Web> InitWebAsync()
        {
            if (web == null)
            {
                web = Context.Web;
                Context.Load(web, t => t.Title);
                await Context.ExecuteQueryAsync();
            }
            return web;
        }

        private void FindAndInitializeSpOnlineLists()
        {
            const BindingFlags bindingFlags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            foreach (var propertyInfo in this.GetType()
                .GetProperties(bindingFlags)
                .Where(t => t.GetIndexParameters().Length == 0
                        && t.DeclaringType != typeof(SharePointContext)))
            {
                var setter = propertyInfo.GetSetMethod();
                var entityConstructor = propertyInfo.PropertyType.GetConstructor(new Type[] { typeof(SharePointContext) });
                var entity = entityConstructor.Invoke(new object[] { this });
                setter.Invoke(this, new object[] { entity });
            }
        }

        public void Dispose()
        {
            Context.Dispose();
            authManager.Dispose();
        }
    }
}
