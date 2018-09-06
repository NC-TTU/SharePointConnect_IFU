///<summary>
/// Erstellt von TTU
/// Erstellt am 12.01.18
/// 
/// 05.04.18 Anpassungen für Nutzung von IFU Content Types. 
/// 
/// Die FolderCreator-Klasse soll auf gerufen werden, wenn
/// ein neuer Kunde in NAV erstellt wird. Dann wird auf dem 
/// SharePoint direkt ein Ordner mit dem Namen des Kunden erstellt.
/// Außerdem werden Unterordner für z.B. Rechnungen angelegt.
/// </summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using log4net;

namespace SharePointConnect
{

    public class FolderCreator {

        private static readonly ILog logger = LogManager.GetLogger(typeof(FolderCreator));

        private List<string> folderNames;
        private ClientContext clientContext;
        private Web site;
        private string rootName;
        private string listUrl;
        private string baseUrl;
        private string subWebsite;
        private string user;
        private string password;
        private string subPage;
        private string listName;



        /**************Konstruktoren*************/

        public FolderCreator() {
            this.folderNames = new List<string>();
        }

        public FolderCreator(string baseUrl, string subWebsite, string user, string password, string subPage, string listName) {
            if (String.IsNullOrEmpty(subPage)) {
                this.listUrl = baseUrl + subWebsite + "/" + listName;
            } else {
                this.listUrl = baseUrl + subWebsite + subPage + "/" + listName;
            }
            this.listName = listName;
            try {
                if (String.IsNullOrEmpty(subPage)) {
                    Connector connector = Connector.GetConnector(baseUrl + subWebsite, user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {
                    Connector connector = Connector.GetConnector(baseUrl + subWebsite + subPage + "/", user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug("Url: " + baseUrl + subWebsite + " SubPage: " + subPage + " User: " + user);
            }
            this.folderNames = new List<string>();
        }


        public void CreateFoldersInSharePoint() {
            if (!HasConnection()) {
                GetConnection();
            }
            Folder parentFolder = null;
            List list = site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();

            list.EnableFolderCreation = true;
            list.Update();
            this.clientContext.ExecuteQuery();

            ListItemCreationInformation creationInformation = new ListItemCreationInformation {
                UnderlyingObjectType = FileSystemObjectType.Folder,
                LeafName = this.rootName
            };

            ListItem newItem = list.AddItem(creationInformation);
            newItem["Title"] = this.rootName;
            newItem.Update();
            this.clientContext.ExecuteQuery();

            list = site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();

            FolderCollection allFolders = list.RootFolder.Folders;
            this.clientContext.Load(allFolders);
            this.clientContext.ExecuteQuery();

            foreach (Folder folder in allFolders) {
                if (folder.Name == "Forms") {
                    continue;
                }
                if (folder.Name != this.rootName) {
                    continue;
                }
                parentFolder = folder;
            }

            foreach (string folderName in this.folderNames) {
                this.clientContext.Load(parentFolder);
                this.clientContext.ExecuteQuery();

                parentFolder.Folders.Add(folderName);
                this.clientContext.ExecuteQuery();
            }

            /****Aufräumarbeit****/
            this.clientContext.Dispose();
            this.site = null;
        }

        // CreateIFUEvent erstellt ein DocumentSet mit der Veranstaltungsnummer als Namen
        // direkt unter dem Rootfolder der Liste.
        public void CreateIFUEvent(string eventTemplateNo) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List eventList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = eventList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                ListItemCreationInformation listItem = new ListItemCreationInformation();
                listItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                listItem.LeafName = eventTemplateNo;
                listItem.FolderUrl = rootFolder.ServerRelativeUrl;

                this.clientContext.Load(eventList.ContentTypes);
                this.clientContext.ExecuteQuery();
                ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Veranstaltung").First();

                var ifuEvent = eventList.AddItem(listItem);
                ifuEvent["ContentTypeId"] = contentType.Id;
                ifuEvent["HTML_x0020_File_x0020_Type"] = "SharePoint.DocumentSet";
                ifuEvent["Title"] = eventTemplateNo;
                ifuEvent.Update();

                this.clientContext.ExecuteQuery();

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("EventNo: " + eventTemplateNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        // CreateIFUTown erstellt einen Ordner in der Liste mit dem Content Type 
        // Stadt.
        public void CreateIFUTown(string eventNo, string eventTemplateNo) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List eventList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = eventList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                ListItemCreationInformation listItem = new ListItemCreationInformation();
                listItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                listItem.LeafName = eventNo;
                listItem.FolderUrl = rootFolder.ServerRelativeUrl + "/" + eventTemplateNo;

                this.clientContext.Load(eventList.ContentTypes);
                this.clientContext.ExecuteQuery();
                ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Stadt").First();

                var ifuTown = eventList.AddItem(listItem);
                ifuTown["ContentTypeId"] = contentType.Id;
                ifuTown["HTML_x0020_File_x0020_Type"] = "SharePoint.Folder";
                ifuTown["Title"] = eventNo;
                ifuTown.Update();

                this.clientContext.ExecuteQuery();

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("folderName: " + eventNo + "EventNo: " + eventTemplateNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        // CreateIFUContact erstellt ein DocumentSet in der Liste mit dem
        // Content Type IFU Kontakt.
        public void CreateIFUContact(string contactNo) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List contactList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = contactList.RootFolder;
                clientContext.Load(rootFolder, sru => sru.ServerRelativeUrl);
                clientContext.ExecuteQuery();

                ListItemCreationInformation listItem = new ListItemCreationInformation();
                listItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                listItem.LeafName = contactNo;
                listItem.FolderUrl = rootFolder.ServerRelativeUrl;

                this.clientContext.Load(contactList.ContentTypes);
                this.clientContext.ExecuteQuery();
                ContentType contentType = contactList.ContentTypes.Where(ct => ct.Name == "IFU Kontakt").First();

                var ifuContact = contactList.AddItem(listItem);
                ifuContact["ContentTypeId"] = contentType.Id;
                ifuContact["HTML_x0020_File_x0020_Type"] = "SharePoint.DocumentSet";
                ifuContact["Title"] = contactNo;
                ifuContact.Update();

                this.clientContext.ExecuteQuery();

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("ContactNo: " + contactNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }

        }

        // Aktualisiert eine Veranstaltung und setzt die Properties vom Contenttype Veranstaltung
        // in der Liste auf die übergebenen Stringwerte;
        public void UpdateIFUEvent(string templateNo, string description, string eventInfo, string startDate) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List eventList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = eventList.RootFolder;
                this.clientContext.Load(rootFolder, sru => sru.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();

                Folder ifuEventFolder = this.site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + templateNo);
                this.clientContext.Load(ifuEventFolder, liaf => liaf.ListItemAllFields);
                this.clientContext.ExecuteQuery();
                var ifuEvent = ifuEventFolder.ListItemAllFields;

                if (ifuEvent != null) {

                    this.clientContext.Load(ifuEvent);
                    this.clientContext.ExecuteQuery();
                    Dictionary<string, object> fieldValues = ifuEvent.FieldValues;
                    if (fieldValues["DocumentSetDescription"] == null) {
                        ifuEvent.ParseAndSetFieldValue("DocumentSetDescription", description);
                    }
                    if (fieldValues["IFUEvent_x002d_Info"] == null) {
                        ifuEvent.ParseAndSetFieldValue("IFUEvent_x002d_Info", eventInfo);
                    }
                    if (fieldValues["IFUStartdate"] == null) {
                        ifuEvent.ParseAndSetFieldValue("IFUStartdate", startDate);
                    }
                    ifuEvent.Update();

                    this.clientContext.ExecuteQuery();

                } else {
                    throw new ArgumentNullException("IFUEvent is null");
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("EventNo: " + templateNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public void UpdateIFUTown(string templateNo, string eventNo, string town, string startDate, string timeFrom, string timeTo, string contactNo) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List eventList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = eventList.RootFolder;
                this.clientContext.Load(rootFolder, sru => sru.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();

                Folder ifuTownFolder = this.site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + templateNo + "/" + eventNo);
                this.clientContext.Load(ifuTownFolder, liaf => liaf.ListItemAllFields);
                this.clientContext.ExecuteQuery();
                var ifuTown = ifuTownFolder.ListItemAllFields;

                if (ifuTown != null) {

                    this.clientContext.Load(ifuTown);
                    this.clientContext.ExecuteQuery();

                    ifuTown.ParseAndSetFieldValue("IFUTown", town);
                    ifuTown.ParseAndSetFieldValue("IFUStartdate", startDate);
                    ifuTown.ParseAndSetFieldValue("IFUTimeFrom", timeFrom);
                    ifuTown.ParseAndSetFieldValue("IFUTimeTo", timeTo);
                    ifuTown.ParseAndSetFieldValue("IFUContactNumber", contactNo);
                    ifuTown.Update();

                    clientContext.ExecuteQuery();

                } else {
                    throw new ArgumentNullException("IFUTown is null");
                }

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("folderName: " + templateNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        // Eine Funktion, die überprüft ob die Eigenschaften eines Kontakts überprüft und true zurückgibt,
        // falls eine der Eigentschaften keinen Wert hat.
        public bool CheckProperties(string contactNo, bool isPerson) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List contactList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = contactList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                Folder ifuContactFolder = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + contactNo);
                clientContext.Load(ifuContactFolder, liaf => liaf.ListItemAllFields);
                clientContext.ExecuteQuery();

                var ifuContact = ifuContactFolder.ListItemAllFields;

                if (ifuContact != null) {

                    if (isPerson) {
                        if (ifuContact.FieldValues["IFUFirm"] == null || String.IsNullOrEmpty(ifuContact.FieldValues["IFUFirm"].ToString())) {
                            return true;
                        }
                        if (ifuContact.FieldValues["IFUFirstname"] == null || String.IsNullOrEmpty(ifuContact.FieldValues["IFUFirstname"].ToString())) {
                            return true;
                        }
                        if (ifuContact.FieldValues["IFUSurname"] == null || String.IsNullOrEmpty(ifuContact.FieldValues["IFUSurname"].ToString())) {
                            return true;
                        }

                    } else {
                        if (ifuContact.FieldValues["IFUFirm"] == null || String.IsNullOrEmpty(ifuContact.FieldValues["IFUFirm"].ToString())) {
                            return true;
                        }
                    }
                    return false; // Wenn keine der Properties von oben null oder Leer ist wird ein false zurückgegeben
                } else {
                    throw new ArgumentNullException("IFU Contact is null");
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("ContactNo: " + contactNo + "Listname: " + this.listName);
                return false;
            }
        }


        public void UpdateIFUContact(string contactNo, string company, string firstname, string surname, bool referee, string title) {
            try {
                if (!HasConnection()) {
                    GetConnection();
                }
                List contactList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = contactList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                Folder ifuContactFolder = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + contactNo);
                clientContext.Load(ifuContactFolder, liaf => liaf.ListItemAllFields);
                clientContext.ExecuteQuery();

                var ifuContact = ifuContactFolder.ListItemAllFields;

                if (ifuContact != null) {

                    this.clientContext.Load(ifuContact);
                    this.clientContext.ExecuteQuery();

                    ifuContact.ParseAndSetFieldValue("IFUFirm", company);
                    ifuContact.ParseAndSetFieldValue("IFUFirstname", firstname);
                    ifuContact.ParseAndSetFieldValue("IFUSurname", surname);
                    ifuContact.ParseAndSetFieldValue("IFUReferee", referee.ToString());
                    ifuContact.ParseAndSetFieldValue("IFUTitle", title);
                    ifuContact.Update();

                    clientContext.ExecuteQuery();

                } else {
                    throw new ArgumentNullException("IFU Contact is null");
                }

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                logger.Debug("ContactNo: " + contactNo + "Listname: " + this.listName);
            } finally {
                /****Aufräumarbeit****/
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public void RenameRootFolder(string oldName, string newName) {
            if (!HasConnection()) {
                GetConnection();
            }
            Folder rootFolder = null;
            if (oldName != newName) {

                try {
                    List list = site.Lists.GetByTitle(this.listName);
                    this.clientContext.Load(list);
                    this.clientContext.ExecuteQuery();

                    FolderCollection allFolders = list.RootFolder.Folders;
                    this.clientContext.Load(allFolders);
                    this.clientContext.ExecuteQuery();

                    foreach (Folder folder in allFolders) {

                        if (folder.Name == "Forms") {
                            continue;
                        }
                        if (folder.Name != this.rootName) {
                            continue;
                        }
                        rootFolder = folder;
                    }

                } catch (System.IO.FileNotFoundException ex) {
                    logger.Error(ex.Message);
                    logger.Debug("RelativeUrl: " + this.listUrl + "/" + oldName + " Oldname:" + oldName + " newName: " + newName);
                    return;
                } finally {
                    /****Aufräumarbeit****/
                    this.clientContext.Dispose();
                    this.site = null;
                }

                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                ListItem folderItem = rootFolder.ListItemAllFields;
                folderItem["Title"] = newName;
                folderItem["FileLeafRef"] = newName;
                folderItem.Update();
                this.clientContext.ExecuteQuery();

            }
        }

        public bool CheckIfFolderAlreadyExists(string folderName) {
            if (!HasConnection()) {
                GetConnection();
            }
            List list = this.site.Lists.GetByTitle(this.listName);
            Folder rootFolder = list.RootFolder;
            this.clientContext.Load(rootFolder, sru => sru.ServerRelativeUrl);
            this.clientContext.ExecuteQuery();

            Folder target = this.site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + folderName);
            try {
                this.clientContext.Load(target);
                this.clientContext.ExecuteQuery();
                return true;
            } catch (ServerException sex) {
                if (sex.ServerErrorTypeName == "System.IO.FileNotFoundException") {
                    target = null;
                    return false;
                }
                logger.Error(sex.Message);
                logger.Debug(sex.StackTrace);
                return true; // wenn irgendwas anderes schief geht tun wir so als ob der Ordner schon da ist.
            }
        }

        public void GetConnection(string baseUrl, string subWebsite, string user, string password, string subPage, string listName) {
    
            if (String.IsNullOrEmpty(subPage)) {
                this.listUrl = baseUrl + subWebsite + "/" + listName;
            } else {
                this.listUrl = baseUrl + subWebsite + subPage + "/" + listName;
            }
            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.user = user;
            this.password = password;
            this.subPage = subPage;
            this.listName = listName;

            try {
                Connector connector = Connector.GetConnector(baseUrl + subWebsite + subPage + "/", user, password);
                this.clientContext = connector.GetClientContext();
                this.site = connector.GetWebSite();
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug("Url: " + baseUrl + subWebsite + " SubPage: " + subPage + " User: " + user);
            }
        }

        private void GetConnection(){
            GetConnection(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName);
        }

        private bool HasConnection() {
            if (this.clientContext == null || this.site == null) {
                return false;
            } else {
                return true;
            }
        }

        public void AddFolderName(string name) { this.folderNames.Add(name); }

        public void AddFolderNames(ICollection<string> collNames) { this.folderNames.AddRange(collNames); }

        public void SetRootName(string rootName) { this.rootName = rootName; }

    }
}