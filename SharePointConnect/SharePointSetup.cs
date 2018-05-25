///<summary>
/// 
/// Die SharePointSetup-Klasse erstellt die Subwebsites,
/// falls diese nicht vorhanden sind.
/// Es kann zu einer DTD Fehlermeldung kommen, wenn dem so 
/// ist muss man in der "hosts.txt" Datei unter "c:\windows\system32\drivers\etc\hosts"
/// folgende User unter localhost anlegen:
/// 
///  	127.0.0.1   msoid.<DOMAIN>.onmicrosoft.com
/// 	127.0.0.1   msoid.onmicrosoft.com
/// 
///</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.Online.SharePoint.TenantAdministration;
using log4net;
using System.Net;

namespace SharePointConnect
{

    public class SharePointSetup
    {

        private static readonly ILog logger = LogManager.GetLogger(typeof(SharePointSetup));

        // Überprüft ob die SubWebsite schon existiert.
        // Es wird eine Anfrage an die Seite geschickt, falls es zu einer Exeption kommt,
        // werden zwei Fälle überprüft. Zum einen ob der Statuscode NotFound 404 ist,
        // das bedeutet die Seite existiert nicht. Im zweiten Fall kann es sein, das die 
        // Seite existiert, aber wir keinen Zugriff haben, deswegen kann es zum Statuscode Forbidden 403
        // kommen. Sollte irgendein andere WebExeption geworfen werden, geht das Programm davon aus, dass
        // die Seite existiert und erstellt keine neue.
        public static bool CheckIfAlreadyExists(string siteName, string baseUrl) {
            Console.WriteLine(baseUrl);
            HttpWebResponse response = null;
            string url = "";
            if (String.IsNullOrEmpty(siteName)) {
                url = baseUrl;
            } else {
                url = baseUrl + siteName ;
            }
            var request = (HttpWebRequest)WebRequest.Create(url);
            Console.WriteLine(url);
            request.Method = "HEAD";
            try {
                response = (HttpWebResponse)request.GetResponse();
                return true;
            } catch (WebException wex) {
                if (((HttpWebResponse)wex.Response).StatusCode == HttpStatusCode.NotFound) {
                    return false;
                } else if (((HttpWebResponse)wex.Response).StatusCode == HttpStatusCode.Forbidden) {
                    return true;
                }
                return true;
            } finally {
                if (response != null) {
                    response.Close();
                }
            }
        }


        public bool StartCreating(string siteName, string baseUrl, string user, string password) {
            try {
                return CreateSubPage(siteName, baseUrl, user, password);
            } catch (Exception ex) {
                logger.Debug(ex.Message);
                logger.Info(ex.StackTrace);
                return false;
            }
        }

        // Diese Funktion verbindet sich mit dem SharePoint und versucht eine neue SubWebsite zu erstellen.
        // Der hinterlegte User sollte Zugriff auf die Administrationsseite des SharePoints haben.
        // https://<DOMAIN>-admin.sharepoint.com/
        private bool CreateSubPage(string siteName, string baseUrl, string user, string password) {

            string tenantUrl = baseUrl.Replace(".sharepoint", "-admin.sharepoint");
            string siteUrl = baseUrl + "Bereich/" + siteName;
            logger.Info(siteUrl);
            logger.Info(tenantUrl);
            Connector connector = Connector.GetConnector(tenantUrl, user, password);
            using (ClientContext context = connector.GetClientContext()) {

                var tenant = new Tenant(context);

                var siteCreationProperties = new SiteCreationProperties() {
                    StorageMaximumLevel = 300,
                    Url = siteUrl,
                    Title = siteName,
                    Owner = user,
                    Template = "STS#0" // Template für Teamsite               
                };

                SpoOperation spo = tenant.CreateSite(siteCreationProperties);

                context.Load(tenant);

                context.Load(spo, i => i.IsComplete);
                context.ExecuteQuery();

                while (!spo.IsComplete) {
                    System.Threading.Thread.Sleep(30000);
                    spo.RefreshLoad();
                    context.ExecuteQuery();
                }
                return true;
            }
        }
    }
}

