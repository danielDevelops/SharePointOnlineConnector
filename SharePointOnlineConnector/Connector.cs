using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public class Connector : IDisposable
    {
        private Web web;
        private readonly AuthenticationManager authManager;
        internal ClientContext Context { get; }

        public Connector(string adAppId, string url, string username, string password)
        {
            authManager = new AuthenticationManager(adAppId);
            var uri = new Uri($"{url}");
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            Context = authManager.GetContext(uri, username, securePassword);
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

        public void Dispose()
        {
            Context.Dispose();
            authManager.Dispose();
        }
    }
}
