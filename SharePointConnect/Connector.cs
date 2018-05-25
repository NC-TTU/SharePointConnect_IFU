///<summary>
/// Erstellt von: NC-TTU 
/// Erstellt am: 29.11.201
/// 
/// Die Connector-Klasse soll die Verbindung zum SharePoint aufbauen und 
/// der DocumentGetter-Klasse ermöglichen über den ClientContext mit dem  
/// SharePoint zu kommunizieren. Die Connector-Klasse ist ein Singleton, 
/// demnach wird es nur eine Instanz der Connector-Klasse geben.
/// </summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Net;

namespace SharePointConnect
{

    class Connector
    {

        public static Connector connector;
        private ClientContext clientContext;
        private Web site;

        public Connector(string baseUrl, string user, string password) {

            SetClientContext(baseUrl, @user, password);
            SetWebSite();
        }

      
        public static Connector GetConnector(string baseUrl, string user, string password) {

         
                connector = new Connector(baseUrl, @user, password);
                return connector;

        }

        public SecureString GetSecureString(string password) {
            // Für die SharePointOnlineCredentials wird ein SecureString aus Microsoft.Security benötigt.
            // GetSecureString nimmt einen String an und läuft über jeden Char in dem String und fügt diesen in dem SecurString hinzu.

            SecureString secure = new SecureString();

            foreach (char c in password) {
                secure.AppendChar(c);
            }

            return secure;
        }


        /*************Getter/Setter************/

        public ClientContext GetClientContext() { return this.clientContext; }

        public Web GetWebSite() { return this.site; }

        public static void Disconnect() { connector = null; }

        private void SetClientContext(string baseUrl, string user, string password) {
            if (user.Contains("\\")) {
                string[] parts = user.Split('\\');
                string domain = parts[0];
                string userName = parts[1];
                this.clientContext = new ClientContext(new Uri(baseUrl)) {
                    Credentials = new NetworkCredential(userName, this.GetSecureString(password), domain)
                };
            } else {
                this.clientContext = new ClientContext(new Uri(baseUrl)) {
                    Credentials = new SharePointOnlineCredentials(user, this.GetSecureString(password))
                };
            }
        }

        private void SetWebSite() { this.site = this.clientContext.Web; }
    }
}
