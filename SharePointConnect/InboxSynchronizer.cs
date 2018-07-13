using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Net;
using System.IO;
using log4net;


namespace SharePointConnect
{
    public static class InboxSynchronizer
    {

        private static readonly ILog logger = LogManager.GetLogger(typeof(InboxSynchronizer));

        #region Registrations
        static public void UpdateRegistration(string baseUrl, string subWebsite, string user, string password, string listname, BrochureOrder brochureOrder) {
            try {
                string[] parts = user.Split('\\');

                using (ClientContext clientContext = new ClientContext(baseUrl + subWebsite)) {
                    clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                    Web site = clientContext.Web;

                    List list = site.Lists.GetByTitle(listname);

                    CamlQuery query = new CamlQuery() {
                        ViewXml = "<View Scope=\"RecursiveAll\">" +
                                        "<Query>" +
                                            "<Where>" +
                                                "<Eq>" +
                                                    "<FieldRef Name=\"GUID\"/>" +
                                                    "<Value Type=\"GUID\">" + brochureOrder.GetGuid() + "</Value>" +
                                                "</Eq>" +
                                            "</Where>" +
                                        "</Query>" +
                                    "</View>"
                    };

                    string status = "";

                    switch (brochureOrder.GetStatus()) {
                        case 0:
                            status = "Neu";
                            break;
                        case 1:
                            status = "Verarbeitet";
                            break;
                        case 2:
                            status = "Zurückgestellt";
                            break;
                    }


                    ListItemCollection itemColl = list.GetItems(query);
                    clientContext.Load(itemColl);
                    clientContext.ExecuteQuery();
                    FieldUrlValue urlValue = new FieldUrlValue();

                    ListItem item = itemColl.First();
                    item.ParseAndSetFieldValue("IFUKuAnmldgStatus", status);
                    item.ParseAndSetFieldValue("IFUKuAnmldgContactNumber", brochureOrder.GetContactNo());
                    item.ParseAndSetFieldValue("IFUEventnumber", brochureOrder.GetEventNo());
                    item.ParseAndSetFieldValue("IFUEventTemplateNumber", brochureOrder.GetEventTemplateNo());
                    item.ParseAndSetFieldValue("IFUCourseNumber", brochureOrder.GetCourseNo());
                    item.ParseAndSetFieldValue("IFUArticleNumber", brochureOrder.GetArticleNo());
                    item.Update();
                    clientContext.ExecuteQuery();
                }
            }catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }

        static public List<string> SplitContactString(string contacts) {
            List<string> contactList = new List<string>();

            contactList.AddRange(contacts.Split('|'));

            return contactList;
        }

        static public void UpdateRegistration(string baseUrl, string subWebsite, string user, string password, string listname, Registration registration) {
            try {
                string[] parts = user.Split('\\');

                using (ClientContext clientContext = new ClientContext(baseUrl + subWebsite)) {
                    clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                    Web site = clientContext.Web;

                    List list = site.Lists.GetByTitle(listname);

                    CamlQuery query = new CamlQuery() {
                        ViewXml = "<View Scope=\"RecursiveAll\">" +
                                      "<Query>" +
                                          "<Where>" +
                                              "<Eq>" +
                                                  "<FieldRef Name=\"GUID\"/>" +
                                                  "<Value Type=\"GUID\">" + registration.GetGuid() + "</Value>" +
                                              "</Eq>" +
                                          "</Where>" +
                                      "</Query>" +
                                  "</View>"
                    };

                    string status = "";

                    switch (registration.GetStatus()) {
                        case 0:
                            status = "Neu";
                            break;
                        case 1:
                            status = "Verarbeitet";
                            break;
                        case 2:
                            status = "Zurückgestellt";
                            break;
                    }


                    ListItemCollection itemColl = list.GetItems(query);
                    clientContext.Load(itemColl);
                    clientContext.ExecuteQuery();
                    FieldUrlValue urlValue = new FieldUrlValue();

                    ListItem item = itemColl.First();
                    item.ParseAndSetFieldValue("IFUKuAnmldgStatus", status);
                    item.ParseAndSetFieldValue("IFUKuAnmldgContactNumber", registration.GetContactNo());
                    item.ParseAndSetFieldValue("IFUEventnumber", registration.GetEventNo());
                    item.ParseAndSetFieldValue("IFUEventTemplateNumber", registration.GetEventTemplateNo());
                    item.ParseAndSetFieldValue("IFUCourseNumber", registration.GetCourseNo());
                    item.ParseAndSetFieldValue("IFUArticleNumber", registration.GetArticleNo());
                    item.Update();
                    clientContext.ExecuteQuery();

                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }
        


        public static List<Registration> SynchronizeRegistrations(string baseUrl, string subWebsite, string user, string password, string listname, string navTempPath) {

            List<Registration> registrations = new List<Registration>();
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(baseUrl + subWebsite)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"IFUKuAnmldgStatus\" />" +
                                              "<Value Type=\"Text\">" + "Neu" + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                foreach (ListItem li in itemColl) {
                    clientContext.Load(li, i => i.File);
                    clientContext.ExecuteQuery();
                    string fileServerRelativePath = li.File.ServerRelativeUrl;
                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileServerRelativePath);
                    string filePath = Path.Combine(navTempPath, li.File.Name);

                    using (FileStream fs = System.IO.File.Create(filePath)) {
                        fileInfo.Stream.CopyTo(fs);
                    }

                    registrations.Add(new Registration(li.FieldValues, filePath));

                }

            }
            return registrations;
        }
        #endregion

        #region Brochure
        public static List<BrochureOrder> SynchronizeOrders(string baseUrl, string subWebsite, string user, string password, string listname, string navTempPath) {

            List<BrochureOrder> brochureOrders = new List<BrochureOrder>();
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(baseUrl + subWebsite)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"IFUKuAnmldgStatus\" />" +
                                              "<Value Type=\"Text\">" + "Neu" + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                foreach (ListItem li in itemColl) {
                    clientContext.Load(li, i => i.File);
                    clientContext.ExecuteQuery();
                    string fileServerRelativePath = li.File.ServerRelativeUrl;
                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileServerRelativePath);
                    string filePath = Path.Combine(navTempPath, li.File.Name);

                    using (FileStream fs = System.IO.File.Create(filePath)) {
                        fileInfo.Stream.CopyTo(fs);
                    }
                    brochureOrders.Add(new BrochureOrder(li.FieldValues, filePath));
                }
            }
            return brochureOrders;
        }

        static public void UpdateBrochureOrder(string baseUrl, string subWebsite, string user, string password, string listname, BrochureOrder brochureOrder) {
            try { 
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(baseUrl + subWebsite)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\">" +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"GUID\"/>" +
                                              "<Value Type=\"GUID\">" + brochureOrder.GetGuid() + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                string status = "";

                switch (brochureOrder.GetStatus()) {
                    case 0:
                        status = "Neu";
                        break;
                    case 1:
                        status = "Verarbeitet";
                        break;
                    case 2:
                        status = "Zurückgestellt";
                        break;
                }


                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();
                FieldUrlValue urlValue = new FieldUrlValue();

                ListItem item = itemColl.First();
                item.ParseAndSetFieldValue("IFUKuAnmldgStatus", status);
                item.ParseAndSetFieldValue("IFUKuAnmldgContactNumber", brochureOrder.GetContactNo());
                item.ParseAndSetFieldValue("IFUEventnumber", brochureOrder.GetEventNo());
                item.ParseAndSetFieldValue("IFUEventTemplateNumber", brochureOrder.GetEventTemplateNo());
                item.ParseAndSetFieldValue("IFUCourseNumber", brochureOrder.GetCourseNo());
                item.ParseAndSetFieldValue("IFUArticleNumber", brochureOrder.GetArticleNo());
                item.Update();
                clientContext.ExecuteQuery();
            }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }

        #endregion

        #region Fee
        public static List<FeeAttachment> GetFeeAttachment(string url, string user, string password, string listname, Guid guid) {
            List<FeeAttachment> feeAttachments = new List<FeeAttachment>();
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\">" +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"GUID\"/>" +
                                              "<Value Type=\"GUID\">" + guid + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                ListItem item = itemColl.First();
                AttachmentCollection attachments = item.AttachmentFiles;
                clientContext.Load(attachments);
                clientContext.ExecuteQuery();
                foreach (Attachment a in attachments) {
                    feeAttachments.Add(new FeeAttachment(url.Substring(0, url.Length - 1) + a.ServerRelativeUrl));
                }
            }

            return feeAttachments;
        }


        public static void SetFileUrlFee(string url, string user, string password, string listname, Guid guid, string serverFilePath) {
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\">" +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"GUID\"/>" +
                                              "<Value Type=\"GUID\">" + guid + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();
                FieldUrlValue urlValue = new FieldUrlValue();
                urlValue.Url = serverFilePath;
                urlValue.Description = "Quittung öffnen";
                ListItem item = itemColl.First();

                item["Receipt"] = urlValue;
                item.ParseAndSetFieldValue("ApprovalState", "PDF Quittung übergeben");
                item.Update();
                clientContext.ExecuteQuery();
            }
        }

        public static void ApproveFee(string url, string user, string password, string listname, Guid guid) {
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\">" +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"GUID\"/>" +
                                              "<Value Type=\"GUID\">" + guid + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                ListItem item = itemColl.First();
                item.ParseAndSetFieldValue("ApprovalState", "genehmigt");
                item.Update();

                clientContext.ExecuteQuery();
            }
        }

        public static void DeclineFee(string url, string user, string password, string listname, Guid guid) {
            string[] parts = user.Split('\\');

            using (ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\">" +
                                  "<Query>" +
                                      "<Where>" +
                                          "<Eq>" +
                                              "<FieldRef Name=\"GUID\"/>" +
                                              "<Value Type=\"GUID\">" + guid + "</Value>" +
                                          "</Eq>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                ListItem item = itemColl.First();
                item.ParseAndSetFieldValue("ApprovalState", "abgehlehnt");
                item.Update();

                clientContext.ExecuteQuery();
            }
        }

        public static List<Fee> SynchronizeFee(string url, string user, string password, string listname) {

            List<Fee> feeList = new List<Fee>();
            string[] parts = user.Split('\\');
            bool hasAttachment = false;

            using (ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                CamlQuery query = new CamlQuery() {
                    ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                  "<Query>" +
                                      "<Where>" +
                                          "<IsNull>" +
                                              "<FieldRef Name=\"ApprovalState\" />" +
                                          "</IsNull>" +
                                      "</Where>" +
                                  "</Query>" +
                              "</View>"
                };

                ListItemCollection itemColl = list.GetItems(query);
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                foreach (ListItem li in itemColl) {
                    clientContext.Load(li.AttachmentFiles);
                    clientContext.ExecuteQuery();

                    if (li.AttachmentFiles.Count > 0) {
                        hasAttachment = true;
                    }


                    feeList.Add(new Fee(li.FieldValues, hasAttachment));
                    li.ParseAndSetFieldValue("ApprovalState", "in Prüfung");
                    li.Update();
                }

                clientContext.ExecuteQuery();

                return feeList;
            }

        }
        #endregion

        #region ScannedDocuments

        public static List<ScannedDocument> SynchronizeScannedDocuments(string url, string user, string password, string listname, string navTempPath) {
            List<ScannedDocument> scannedDocuments = new List<ScannedDocument>();
            string[] parts = user.Split('\\');

            using(ClientContext clientContext = new ClientContext(url)) {
                clientContext.Credentials = new NetworkCredential(parts[1], password, parts[0]);

                Web site = clientContext.Web;

                List list = site.Lists.GetByTitle(listname);

                ListItemCollection itemColl = list.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(itemColl);
                clientContext.ExecuteQuery();

                for(int i = 0; i < itemColl.Count; ++i){
                    ListItem li = itemColl[i];
                    clientContext.Load(li, f => f.File);
                    clientContext.ExecuteQuery();
                    string fileServerRelativePath = li.File.ServerRelativeUrl;
                    FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileServerRelativePath);
                    string filePath = Path.Combine(navTempPath, li.File.Name);

                    using (FileStream fs = System.IO.File.Create(filePath)) {
                        fileInfo.Stream.CopyTo(fs);
                    }
                    scannedDocuments.Add(new ScannedDocument(li.FieldValues, filePath));

                    li.DeleteObject();                 
                }
                clientContext.ExecuteQuery();
            }
            return scannedDocuments;
        } 

        #endregion
    }
}
