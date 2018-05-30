///<summary>
///  Erstellt von: NC-TTU                                                  
///  Erstellt am: 22.12.2017  
/// 
/// 
/// 
///  Letzte Änderung 15.03.18
///  Die Funktion UploadFileFromNav wurde implementiert, sie nimmt 
///  einen Dateipfad und ein Dokumententyp, sucht dann den richtigen
///  Ordner im SharePoint und übergibt den Pfad und den Ordner an
///  UploadFile.
///  
///  Die SharePointFileUploader-Klasse wird vom SharePointUserControl      
///  aufgerufen und erhält die Verbindungsdaten um eine Verbindung über
///  die Connector-Klasse mit dem SharePoint aufzubauen. Danach wird von
///  der SharePointUserControl-Klasse die Funktion UploadFile mit dem
///  Elternordner der Datei und dem Dateipfad aufgerufen.
/// </summary>

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using log4net;
using System.Collections;
using System.Globalization;
using System.Resources;
using System.Reflection;
using Microsoft.SharePoint.Client.Taxonomy;

namespace SharePointConnect
{

    public class SharePointFileUploader {

        private static readonly ILog logger = LogManager.GetLogger(typeof(SharePointFileUploader));

        private ClientContext clientContext;
        private Web site;

        private string url;
        private string baseUrl;
        private string subWebsite;
        private string user;
        private string password;
        private string subPage;
        private string listName;
        private string filter;
        private SharePointFileGetter fileGetter;

        /****************Konstruktoren*****************/
        public SharePointFileUploader() { }

        public SharePointFileUploader(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {

            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.user = user;
            this.password = password;
            this.subPage = subPage;
            this.listName = listName;
            this.filter = filter;

            if (String.IsNullOrEmpty(subPage)) {
                this.url = baseUrl + subWebsite;
            } else {
                this.url = baseUrl + subWebsite + subPage + "/";
            }
        }

        public SharePointFileUploader(SharePointFolder parentFolder, string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {

            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.user = user;
            this.password = password;
            this.subPage = subPage;
            this.listName = listName;
            this.filter = filter;

            this.url = this.baseUrl.Substring(0, this.baseUrl.Length - 1);
            this.url += parentFolder.GetServerRelativeUrl();
            Connector connector = Connector.GetConnector(this.url, user, password);
            this.clientContext = connector.GetClientContext();
            this.site = connector.GetWebSite();

        }


        // UploadFile bekommt den Dateipfad von der Datei, die hochgeladen werden soll und den 
        // Elternordner der Datei. Danach wird die FileCreationInformation-Klasse mit den Datei
        // Informationen gefüllt und am Ende zu einer SharePoint.Client.File gemacht.
        // Und dann wird diese durch den ClientContext hochgeladen.
        public void UploadFile(string filePath, SharePointFolder parentFolder) {
            try {
                FileCreationInformation newFile = new FileCreationInformation();
                Folder folder = this.site.GetFolderByServerRelativeUrl(parentFolder.GetServerRelativeUrl());
                Microsoft.SharePoint.Client.File uploadFile;

                this.clientContext.Load(folder);
                this.clientContext.ExecuteQuery();


                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length <= 1200000) { // wenn die Datei kleiner als 1.2 MB ist soll sie direkt hochgeladen werden.
                    newFile.Content = System.IO.File.ReadAllBytes(filePath);
                    newFile.Url = " /" + fileInfo.Name;
                    newFile.Overwrite = true;


                    if (CheckIfFileAlreadyExists(filePath, folder)) {

                        Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(parentFolder.GetServerRelativeUrl() + "/" + fileInfo.Name);
                        this.clientContext.Load(file);
                        this.clientContext.ExecuteQuery();
                        file.CheckOut();

                        uploadFile = folder.Files.Add(newFile);
                        uploadFile.CheckIn("", CheckinType.MajorCheckIn);
                        this.clientContext.Load(uploadFile);
                        this.clientContext.ExecuteQuery();
                    } else {

                        uploadFile = folder.Files.Add(newFile);
                        this.clientContext.Load(uploadFile);
                        this.clientContext.ExecuteQuery();
                    }

                } else {// kann größere Dateien hochladen.
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {
                        FileCreationInformation NewLargeFile = new FileCreationInformation {
                            ContentStream = fs,
                            Url = Path.GetFileName(filePath),
                            Overwrite = true
                        };

                        if (CheckIfFileAlreadyExists(filePath, folder)) {

                            Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(parentFolder.GetServerRelativeUrl() + "/" + fileInfo.Name);
                            this.clientContext.Load(file);
                            this.clientContext.ExecuteQuery();
                            file.CheckOut();

                            uploadFile = folder.Files.Add(NewLargeFile);
                            uploadFile.CheckIn("", CheckinType.MajorCheckIn);
                            this.clientContext.Load(uploadFile);
                            this.clientContext.ExecuteQuery();
                        } else {

                            uploadFile = folder.Files.Add(NewLargeFile);
                            this.clientContext.Load(uploadFile);
                            this.clientContext.ExecuteQuery();
                        }
                    }
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public void UploadFile(string filePath, string folderServerRelativeUrl) {


            FileCreationInformation newFile = new FileCreationInformation();
            Folder folder = this.site.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
            Microsoft.SharePoint.Client.File uploadFile;

            this.clientContext.Load(folder);
            this.clientContext.ExecuteQuery();


            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length <= 1200000) { // wenn die Datei kleiner als 1.2 MB ist soll sie direkt hochgeladen werden.
                newFile.Content = System.IO.File.ReadAllBytes(filePath);
                newFile.Url = " /" + fileInfo.Name;
                newFile.Overwrite = true;


                if (CheckIfFileAlreadyExists(filePath, folder)) {

                    Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(folderServerRelativeUrl + "/" + fileInfo.Name);
                    this.clientContext.Load(file);
                    this.clientContext.ExecuteQuery();
                    file.CheckOut();

                    uploadFile = folder.Files.Add(newFile);
                    uploadFile.CheckIn("", CheckinType.MajorCheckIn);
                    this.clientContext.Load(uploadFile);
                    this.clientContext.ExecuteQuery();
                } else {

                    uploadFile = folder.Files.Add(newFile);
                    this.clientContext.Load(uploadFile);
                    this.clientContext.ExecuteQuery();
                }
            } else {// kann größere Dateien hochladen.
                using (FileStream fs = new FileStream(filePath, FileMode.Open)) {
                    FileCreationInformation NewLargeFile = new FileCreationInformation {
                        ContentStream = fs,
                        Url = Path.GetFileName(filePath),
                        Overwrite = true
                    };

                    if (CheckIfFileAlreadyExists(filePath, folder)) {

                        Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(folderServerRelativeUrl + "/" + fileInfo.Name);
                        this.clientContext.Load(file);
                        this.clientContext.ExecuteQuery();
                        file.CheckOut();

                        uploadFile = folder.Files.Add(NewLargeFile);
                        uploadFile.CheckIn("", CheckinType.MajorCheckIn);
                        this.clientContext.Load(uploadFile);
                        this.clientContext.ExecuteQuery();
                    } else {

                        uploadFile = folder.Files.Add(NewLargeFile);
                        this.clientContext.Load(uploadFile);
                        this.clientContext.ExecuteQuery();
                    }
                }
            }

            /****Aufräumarbeit****/
            Connector.Disconnect();
            this.clientContext.Dispose();
            this.site = null;
        }

        // Aktualisiert die SharePointDataList und gibt diese zurück an die
        // SharePointUserControl.
        public ArrayList UpdateSharePointDataList() {

            this.fileGetter = new SharePointFileGetter(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter);
            return this.fileGetter.GetFiles();
        }


        // Diese Funktion versucht für ein IFUDocument alle Managed Metadatafields zu setzen.
        // Es bekommt ein Dictionary, dass den Managed Metadata Term als Value und den Staticname als Key hat übergeben.
        // Sowie die itemID des betroffenen ListItems und den Namen der Datei.
        private void UpdateIFUDocumentMetadata(Dictionary<string, string> termDictionary, string itemID, string fileName) {
            try {
                foreach (KeyValuePair<string, string> pair in termDictionary) {
                    List list = this.site.Lists.GetByTitle(this.listName);
                    FieldCollection fields = list.Fields;
                    Field field = fields.GetByInternalNameOrTitle(pair.Key);

                    this.clientContext.Load(fields);
                    this.clientContext.Load(field);
                    this.clientContext.ExecuteQuery();

                    TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);
                    string termId = GetTermIdForTerm(pair.Value, txField.TermSetId);

                    this.clientContext.Load(list.RootFolder);
                    this.clientContext.ExecuteQuery();
                    Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + fileName);
                    ListItem item = file.ListItemAllFields;
                    this.clientContext.Load(item);
                    this.clientContext.ExecuteQuery();

                    TaxonomyFieldValue termValue = new TaxonomyFieldValue {
                        Label = pair.Value,
                        TermGuid = termId,
                        WssId = -1
                    };

                    txField.SetFieldValueByValue(item, termValue);
                    item.Update();
                    this.clientContext.Load(item);
                    this.clientContext.ExecuteQuery();
                }
            }catch (ServerException sex) {
                logger.Error(sex.Message);
                logger.Debug(sex.StackTrace);
                foreach(KeyValuePair<string, string> pair in termDictionary) {
                    logger.Debug("Key:" + pair.Key + " " + "Value:" + pair.Value);
                }
            }
        }

        // Diese Funktion versucht für ein Event den DocumentType zu setzen. Der Funktion wird der Documenttype übergeben,
        // sowie die ItemID von der Datei, der Name der Datei, die Veranstaltungsnummer und die Nummer der Stadt.
        private void UpdateDocumentTypeEvent(string documentType, string itemID, string fileName, string eventNo, string ifuTown) {
            try {
                List list = this.site.Lists.GetByTitle(this.listName);
                FieldCollection fields = list.Fields;
                Field field = fields.GetByInternalNameOrTitle("IFUDocumenttype"); // IFUDocumenttype ist der Staticname des Feldes im SharePoint.

                this.clientContext.Load(fields);
                this.clientContext.Load(field);
                this.clientContext.ExecuteQuery();

                TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);
                string termId = GetTermIdForTerm(documentType, txField.TermSetId);

                this.clientContext.Load(list.RootFolder);
                this.clientContext.ExecuteQuery(); 
                Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + eventNo + "/" + ifuTown + "/" + fileName);
                ListItem item = file.ListItemAllFields;
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();

                TaxonomyFieldValue termValue = new TaxonomyFieldValue {
                    Label = documentType,
                    TermGuid = termId,
                    WssId = -1
                };

                txField.SetFieldValueByValue(item, termValue);
                item.Update();
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();
            }catch(ServerException sex) {
                logger.Error(sex.Message);
                logger.Debug(sex.StackTrace);
                logger.Debug("Term:" + documentType);
            }
        }

        // Diese Funktion versucht einer Kontaktdatei einen Documenttype zuzuweisen und bekommt den documenttype,
        // die ItemID, den Dateinamen sowie die zugeordnete Kontaktnummer.
        private void UpdateDocumentTypeContact(string documentType, string itemID, string fileName, string contactNo) {
            try {
                List list = this.site.Lists.GetByTitle(this.listName);
                FieldCollection fields = list.Fields;
                Field field = fields.GetByInternalNameOrTitle("IFUDocumenttype");

                this.clientContext.Load(fields);
                this.clientContext.Load(field);
                this.clientContext.ExecuteQuery();

                TaxonomyField txField = this.clientContext.CastTo<TaxonomyField>(field);
                string termId = GetTermIdForTerm(documentType, txField.TermSetId);

                this.clientContext.Load(list.RootFolder);
                this.clientContext.ExecuteQuery();
                Microsoft.SharePoint.Client.File file = this.site.GetFileByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + contactNo + "/" + fileName);
                ListItem item = file.ListItemAllFields;
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();

                TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                termValue.Label = documentType;
                termValue.TermGuid = termId;
                termValue.WssId = -1;

                txField.SetFieldValueByValue(item, termValue);
                item.Update();
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();
            }catch(ServerException sex) {
                logger.Error(sex.Message);
                logger.Debug(sex.StackTrace);
                logger.Debug("Term:" + documentType);
            }
        }

        // GetTermIdFromTerm bekommt einmal den Documenttype und die GUID des Feldes übergeben.
        // Danach versucht die Funktion über die TaxonomySession an erlaubten Metadaten für
        // das Feld zu kommen. Wenn der Documenttype zulässig ist, wird am Ende die Documenttype GUID
        // zurückgegeben.
        private string GetTermIdForTerm(string term, Guid termSetId) {
            string termId = string.Empty;

            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore ts = tSession.GetDefaultSiteCollectionTermStore();
            TermSet tset = ts.GetTermSet(termSetId);

            LabelMatchInformation lmi = new LabelMatchInformation(clientContext) {
                Lcid = 1033,
                TrimUnavailable = true,
                TermLabel = term
            };

            TermCollection termMatches = tset.GetTerms(lmi);
            this.clientContext.Load(tSession);
            this.clientContext.Load(ts);
            this.clientContext.Load(tset);
            this.clientContext.Load(termMatches);

            this.clientContext.ExecuteQuery();

            if (termMatches != null && termMatches.Count() > 0)
                termId = termMatches.First().Id.ToString();

            return termId;

        }
        // UploadIFUDocument bekommt den Dateipfad und 4 ManagedMetadatavalues übergeben und versucht diese dann in den SharePoint
        // hochzuladen. 
        public void UploadIFUDocument(string filePath, string customer, string categorie, string documentType, string businessYear) {
            try {
                FileCreationInformation fileCreation = new FileCreationInformation();
                Microsoft.SharePoint.Client.File file;
                string fileServerRelativePath = String.Empty;
                string itemID = String.Empty;

                List documentList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = documentList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                FileInfo fileInfo = new FileInfo(filePath);

                if (fileInfo.Length <= 1200000) {

                    fileCreation.Content = System.IO.File.ReadAllBytes(filePath);
                    fileCreation.Url = fileInfo.Name;
                    fileCreation.Overwrite = true;

                    if (CheckIfFileAlreadyExists(filePath, rootFolder)) {

                        Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                        this.clientContext.Load(existingFile);
                        this.clientContext.ExecuteQuery();
                        existingFile.CheckOut();

                        this.clientContext.Load(documentList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = documentList.ContentTypes.Where(ct => ct.Name == "Dokument").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    } else {

                        this.clientContext.Load(documentList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = documentList.ContentTypes.Where(ct => ct.Name == "Dokument").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    }
                } else {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {

                        fileCreation.ContentStream = fs;
                        fileCreation.Url = fileInfo.Name;
                        fileCreation.Overwrite = true;

                        if (CheckIfFileAlreadyExists(filePath, rootFolder)) {

                            Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                            this.clientContext.Load(existingFile);
                            this.clientContext.ExecuteQuery();
                            existingFile.CheckOut();

                            this.clientContext.Load(documentList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = documentList.ContentTypes.Where(ct => ct.Name == "Dokument").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        } else {

                            this.clientContext.Load(documentList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = documentList.ContentTypes.Where(ct => ct.Name == "Dokument").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        }
                    }
                }

                file = this.site.GetFileByServerRelativeUrl(fileServerRelativePath);
                ListItem item = file.ListItemAllFields;
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();
                itemID = item.Id.ToString();

                Dictionary<string, string> termDictionary = new Dictionary<string, string>();
                if (!String.IsNullOrEmpty(customer))
                    termDictionary.Add("Customer", customer);
                if (!String.IsNullOrEmpty(categorie))
                    termDictionary.Add("IFUCategory", categorie);
                if (!String.IsNullOrEmpty(documentType))
                    termDictionary.Add("IFUDocumenttype", documentType);
                if (!String.IsNullOrEmpty(businessYear))
                    termDictionary.Add("Fiscalyear", businessYear);

                if (!String.IsNullOrEmpty(itemID) && termDictionary.Count > 0) {
                    UpdateIFUDocumentMetadata(termDictionary, itemID, fileInfo.Name);
                }

            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }
    
        // UploadRegistration lädt ein eingescanntest Dokument aus NAV wieder in den SharePoint hoch und setzt dabei,
        // falls übergeben Veranstaltungsnummer, Veranstaltungsvorlagennummer und Kontaktnummer.
        public void UploadRegistration(string filePath, string eventNo, string templateNo, string contactNo) {
            try {
                FileCreationInformation fileCreation = new FileCreationInformation();

                List registrationList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = registrationList.RootFolder;
                this.clientContext.Load(rootFolder);
                this.clientContext.ExecuteQuery();

                FileInfo fileInfo = new FileInfo(filePath);

                if (fileInfo.Length <= 1200000) {

                    fileCreation.Content = System.IO.File.ReadAllBytes(filePath);
                    fileCreation.Url = fileInfo.Name;
                    fileCreation.Overwrite = true;

                    if (CheckIfFileAlreadyExists(filePath, rootFolder)) {

                        Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                        this.clientContext.Load(existingFile);
                        this.clientContext.ExecuteQuery();
                        existingFile.CheckOut();

                        this.clientContext.Load(registrationList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = registrationList.ContentTypes.Where(ct => ct.Name == "Anmeldung").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["IFUEventnumber"] = eventNo;
                        uploadedFile.ListItemAllFields["IFUEventTemplateNumber"] = templateNo;
                        uploadedFile.ListItemAllFields["IFUKuAnmldgContactNumber"] = contactNo;
                        uploadedFile.ListItemAllFields["IFUKuAnmldgStatus"] = "neu";
                        uploadedFile.ListItemAllFields.Update();
                        uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                        this.clientContext.ExecuteQuery();
                    } else {

                        this.clientContext.Load(registrationList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = registrationList.ContentTypes.Where(ct => ct.Name == "Anmeldung").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["IFUEventnumber"] = eventNo;
                        uploadedFile.ListItemAllFields["IFUEventTemplateNumber"] = templateNo;
                        uploadedFile.ListItemAllFields["IFUKuAnmldgContactNumber"] = contactNo;
                        uploadedFile.ListItemAllFields["IFUKuAnmldgStatus"] = "neu";
                        uploadedFile.ListItemAllFields.Update();

                        this.clientContext.ExecuteQuery();
                    }
                } else {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {

                        fileCreation.ContentStream = fs;
                        fileCreation.Url = fileInfo.Name;
                        fileCreation.Overwrite = true;

                        if (CheckIfFileAlreadyExists(filePath, rootFolder)) {

                            Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                            this.clientContext.Load(existingFile);
                            this.clientContext.ExecuteQuery();
                            existingFile.CheckOut();

                            this.clientContext.Load(registrationList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = registrationList.ContentTypes.Where(ct => ct.Name == "Anmeldung").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["IFUEventnumber"] = eventNo;
                            uploadedFile.ListItemAllFields["IFUEventTemplateNumber"] = templateNo;
                            uploadedFile.ListItemAllFields["IFUKuAnmldgContactNumber"] = contactNo;
                            uploadedFile.ListItemAllFields["IFUKuAnmldgStatus"] = "neu";
                            uploadedFile.ListItemAllFields.Update();
                            uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                            this.clientContext.ExecuteQuery();
                        } else {

                            this.clientContext.Load(registrationList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = registrationList.ContentTypes.Where(ct => ct.Name == "Anmeldung").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["IFUEventnumber"] = eventNo;
                            uploadedFile.ListItemAllFields["IFUEventTemplateNumber"] = templateNo;
                            uploadedFile.ListItemAllFields["IFUKuAnmldgContactNumber"] = contactNo;
                            uploadedFile.ListItemAllFields["IFUKuAnmldgStatus"] = "neu";
                            uploadedFile.ListItemAllFields.Update();

                            this.clientContext.ExecuteQuery();
                        }
                    }
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public void UploadEventDocument(string filePath, string eventNo, string ifuTown, string title, string activityDate, string documentType) {
            try {
                DateTime dt = DateTime.Parse(activityDate);
                FileCreationInformation fileCreation = new FileCreationInformation();
                Microsoft.SharePoint.Client.File file;
                string fileServerRelativePath = String.Empty;
                string itemID = String.Empty;

                List eventList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = eventList.RootFolder;
                this.clientContext.Load(rootFolder, rf => rf.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();

                Folder parent = this.site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + eventNo + "/" + ifuTown);
                this.clientContext.Load(parent);
                this.clientContext.ExecuteQuery();

                FileInfo fileInfo = new FileInfo(filePath);

                if (fileInfo.Length <= 1200000) {

                    fileCreation.Content = System.IO.File.ReadAllBytes(filePath);
                    fileCreation.Url = fileInfo.Name;
                    fileCreation.Overwrite = true;

                    if (CheckIfFileAlreadyExists(filePath, parent)) {

                        Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(parent.ServerRelativeUrl + "/" + fileInfo.Name);
                        this.clientContext.Load(existingFile);
                        this.clientContext.ExecuteQuery();
                        existingFile.CheckOut();

                        this.clientContext.Load(eventList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Veranstaltungsdokument").First();

                        var uploadedFile = parent.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["Title"] = title;
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                        uploadedFile.ListItemAllFields.Update();
                        uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    } else {

                        this.clientContext.Load(eventList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Veranstaltungsdokument").First();

                        var uploadedFile = parent.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["Title"] = title;
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                        uploadedFile.ListItemAllFields.Update();

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    }
                } else {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {

                        fileCreation.ContentStream = fs;
                        fileCreation.Url = fileInfo.Name;
                        fileCreation.Overwrite = true;

                        if (CheckIfFileAlreadyExists(filePath, parent)) {

                            Microsoft.SharePoint.Client.File existingFile = this.site.GetFileByServerRelativeUrl(parent.ServerRelativeUrl + "/" + fileInfo.Name);
                            this.clientContext.Load(existingFile);
                            this.clientContext.ExecuteQuery();
                            existingFile.CheckOut();

                            this.clientContext.Load(eventList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Veranstaltungsdokument").First();

                            var uploadedFile = parent.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["Title"] = title;
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                            uploadedFile.ListItemAllFields.Update();
                            uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        } else {

                            this.clientContext.Load(eventList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = eventList.ContentTypes.Where(ct => ct.Name == "Veranstaltungsdokument").First();

                            var uploadedFile = parent.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["Title"] = title;
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                            uploadedFile.ListItemAllFields.Update();

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        }
                    }
                }


                file = this.site.GetFileByServerRelativeUrl(fileServerRelativePath);
                ListItem item = file.ListItemAllFields;
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();
                itemID = item.Id.ToString();

                if (!String.IsNullOrEmpty(itemID) && !String.IsNullOrEmpty(documentType)) {
                    UpdateDocumentTypeEvent(documentType, itemID, fileInfo.Name, eventNo, ifuTown);
                }
            } catch(Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public void UploadIFUContactDocument(string filePath, string contactNo, string title, string activityDate, string documentType) {
            try {
                DateTime dt = DateTime.Parse(activityDate);
                FileCreationInformation fileCreation = new FileCreationInformation();
                Microsoft.SharePoint.Client.File file;
                string fileServerRelativePath = String.Empty;
                string itemID = String.Empty;

                List contactList = site.Lists.GetByTitle(listName);
                Folder rootFolder = contactList.RootFolder;
                this.clientContext.Load(rootFolder, rf => rf.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();

                Folder parent = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + contactNo);
                this.clientContext.Load(parent);
                this.clientContext.ExecuteQuery();

                FileInfo fileInfo = new FileInfo(filePath);

                if (fileInfo.Length <= 1200000) {

                    fileCreation.Content = System.IO.File.ReadAllBytes(filePath);
                    fileCreation.Url = fileInfo.Name;
                    fileCreation.Overwrite = true;

                    if (CheckIfFileAlreadyExists(filePath, parent)) {

                        Microsoft.SharePoint.Client.File existingFile = site.GetFileByServerRelativeUrl(parent.ServerRelativeUrl + "/" + fileInfo.Name);
                        this.clientContext.Load(existingFile);
                        this.clientContext.ExecuteQuery();
                        existingFile.CheckOut();

                        this.clientContext.Load(contactList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = contactList.ContentTypes.Where(ct => ct.Name == "Kontaktdokument").First();

                        var uploadedFile = parent.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["Title"] = title;
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["IFUContactNumber"] = contactNo;
                        uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                        uploadedFile.ListItemAllFields.Update();
                        uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    } else {

                        this.clientContext.Load(contactList.ContentTypes);
                        this.clientContext.ExecuteQuery();
                        ContentType contentType = contactList.ContentTypes.Where(ct => ct.Name == "Kontaktdokument").First();

                        var uploadedFile = parent.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["Title"] = title;
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["IFUContactNumber"] = contactNo;
                        uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                        uploadedFile.ListItemAllFields.Update();

                        this.clientContext.ExecuteQuery();

                        this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                        this.clientContext.ExecuteQuery();
                        fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                    }
                } else {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {

                        fileCreation.ContentStream = fs;
                        fileCreation.Url = fileInfo.Name;
                        fileCreation.Overwrite = true;

                        if (CheckIfFileAlreadyExists(filePath, parent)) {

                            Microsoft.SharePoint.Client.File existingFile = site.GetFileByServerRelativeUrl(parent.ServerRelativeUrl + "/" + fileInfo.Name);
                            this.clientContext.Load(existingFile);
                            this.clientContext.ExecuteQuery();
                            existingFile.CheckOut();

                            this.clientContext.Load(contactList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = contactList.ContentTypes.Where(ct => ct.Name == "Kontaktdokument").First();

                            var uploadedFile = parent.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["Title"] = title;
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["IFUContactNumber"] = contactNo;
                            uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                            uploadedFile.ListItemAllFields.Update();
                            uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        } else {

                            this.clientContext.Load(contactList.ContentTypes);
                            this.clientContext.ExecuteQuery();
                            ContentType contentType = contactList.ContentTypes.Where(ct => ct.Name == "Kontaktdokument").First();

                            var uploadedFile = parent.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["Title"] = title;
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["IFUContactNumber"] = contactNo;
                            uploadedFile.ListItemAllFields["IFUActivityDate"] = dt;
                            uploadedFile.ListItemAllFields.Update();

                            this.clientContext.ExecuteQuery();

                            this.clientContext.Load(uploadedFile, uf => uf.ServerRelativeUrl);
                            this.clientContext.ExecuteQuery();
                            fileServerRelativePath = uploadedFile.ServerRelativeUrl;
                        }
                    }
                }


                file = this.site.GetFileByServerRelativeUrl(fileServerRelativePath);
                ListItem item = file.ListItemAllFields;
                this.clientContext.Load(item);
                this.clientContext.ExecuteQuery();
                itemID = item.Id.ToString();

                if (!String.IsNullOrEmpty(itemID) && !String.IsNullOrEmpty(documentType)) {
                    UpdateDocumentTypeContact(documentType, itemID, fileInfo.Name, contactNo);
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }

        public string UploadIFUFee(string filePath, string contactNo) {
            try {
                FileCreationInformation fileCreation = new FileCreationInformation();

                string serverFilePath = "";

                List receiptList = this.site.Lists.GetByTitle(this.listName);
                Folder rootFolder = receiptList.RootFolder;
                this.clientContext.Load(rootFolder, rf => rf.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();

                FileInfo fileInfo = new FileInfo(filePath);

                if (fileInfo.Length <= 1200000) {
                    fileCreation.Content = System.IO.File.ReadAllBytes(filePath);
                    fileCreation.Url = fileInfo.Name;
                    fileCreation.Overwrite = true;

                    if (CheckIfFileAlreadyExists(filePath, rootFolder)) {
                        Microsoft.SharePoint.Client.File existingFile = site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                        clientContext.Load(existingFile);
                        clientContext.ExecuteQuery();
                        existingFile.CheckOut();

                        clientContext.Load(receiptList.ContentTypes);
                        clientContext.ExecuteQuery();
                        ContentType contentType = receiptList.ContentTypes.Where(ct => ct.Name == "Quittung").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["Contactnumber"] = contactNo;
                        uploadedFile.ListItemAllFields.Update();
                        uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                        clientContext.ExecuteQuery();
                        serverFilePath = this.baseUrl.Substring(0, this.baseUrl.Length - 1) + rootFolder.ServerRelativeUrl + "/" + fileInfo.Name;
                    } else {

                        clientContext.Load(receiptList.ContentTypes);
                        clientContext.ExecuteQuery();
                        ContentType contentType = receiptList.ContentTypes.Where(ct => ct.Name == "Quittung").First();

                        var uploadedFile = rootFolder.Files.Add(fileCreation);
                        uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                        uploadedFile.ListItemAllFields["Contactnumber"] = contactNo;
                        uploadedFile.ListItemAllFields.Update();

                        clientContext.ExecuteQuery();
                        serverFilePath = this.baseUrl.Substring(0, this.baseUrl.Length - 1) + rootFolder.ServerRelativeUrl + "/" + fileInfo.Name;
                    }
                } else {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open)) {

                        fileCreation.ContentStream = fs;
                        fileCreation.Url = fileInfo.Name;
                        fileCreation.Overwrite = true;

                        if (CheckIfFileAlreadyExists(filePath, rootFolder)) {

                            Microsoft.SharePoint.Client.File existingFile = site.GetFileByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + fileInfo.Name);
                            clientContext.Load(existingFile);
                            clientContext.ExecuteQuery();
                            existingFile.CheckOut();

                            clientContext.Load(receiptList.ContentTypes);
                            clientContext.ExecuteQuery();
                            ContentType contentType = receiptList.ContentTypes.Where(ct => ct.Name == "Quittung").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["Contactnumber"] = contactNo;
                            uploadedFile.ListItemAllFields.Update();
                            uploadedFile.CheckIn("", CheckinType.MajorCheckIn);

                            clientContext.ExecuteQuery();

                            serverFilePath = this.baseUrl.Substring(0, this.baseUrl.Length - 1) + rootFolder.ServerRelativeUrl + "/" + fileInfo.Name;
                        } else {

                            clientContext.Load(receiptList.ContentTypes);
                            clientContext.ExecuteQuery();
                            ContentType contentType = receiptList.ContentTypes.Where(ct => ct.Name == "Quittung").First();

                            var uploadedFile = rootFolder.Files.Add(fileCreation);
                            uploadedFile.ListItemAllFields["ContentTypeId"] = contentType.Id;
                            uploadedFile.ListItemAllFields["Contactnumber"] = contactNo;
                            uploadedFile.ListItemAllFields.Update();

                            clientContext.ExecuteQuery();

                            serverFilePath = this.baseUrl.Substring(0, this.baseUrl.Length - 1) + rootFolder.ServerRelativeUrl + "/" + fileInfo.Name;
                        }
                    }
                }
                return serverFilePath;
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
                FileInfo info = new FileInfo(filePath);
                throw new Exception("Error by uploading file: " + info.Name);
            } finally {
                /****Aufräumarbeit****/
                Connector.Disconnect();
                this.clientContext.Dispose();
                this.site = null;
            }
        }



        private bool CheckIfFileAlreadyExists(string filePath, Folder parent) {

            var fileInfo = new FileInfo(filePath);
            string fileName = fileInfo.Name;

            FileCollection files = parent.Files;

            this.clientContext.Load(files);
            this.clientContext.ExecuteQuery();

            foreach (Microsoft.SharePoint.Client.File file in files) {

                if (fileName == file.Name) {
                    return true;
                }
            }
            return false;
        }

        /****Getter/Setter****/
        public void SetUrl(SharePointFolder parentFolder) {

            this.url = this.baseUrl.Substring(0, this.baseUrl.Length - 1);
            this.url += parentFolder.GetServerRelativeUrl();
        }

        public void SetConnection() {

            Connector connector = Connector.GetConnector(this.url, this.user, this.password);
            this.clientContext = connector.GetClientContext();
            this.site = connector.GetWebSite();
        }
    }
}
