///<summary>
/// Letzte Änderung am 15.12.17 
///     -- Einen zweiten Konstruktor hinzugefügt, der den Filter mit übernimmt. 
/// Änderung am 08.12.2017 
///     -- Umstellung der Klasse zur Nutzung von SharePointFoldern 
///     
/// Erstellt von: NC-TTU
/// Erstellt am: 29.11.2017 
/// 
/// Die SharePointFileGetter-Klasse soll von Microsoft Dynamics NAV die nötigen Parameter bekommen
/// um eine Verbindung mit dem SharePoint aufbauen zu können.
/// Und Documentfiles aus einer bestimmten Liste von SharePoint in NAV anzuzeigen, die dann von NAV geöffnet werden können.
///</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Diagnostics;
using log4net;
using System.Collections;

namespace SharePointConnect
{

    public class SharePointFileGetter
    {

        private static readonly ILog logger = LogManager.GetLogger(typeof(SharePointFileGetter));

        #region Instancevaribles
        private ClientContext clientContext;
        private Web site;
        private Folder filteredFolder;
        private string baseUrl;  // Url der SharePoint-Seite ohne SubWebsite
        private string subWebsite;
        private string url;      // Url der SharePoint-Seite mit SubWebsite
        private string filter;   // Der Ordername auf SharePoint in dem die Dateien sind.
        private string subPage;  // Name der SubPage in SharePoint, in der sich die Ordner befinden
        private string listName; // Name der Liste in SharePoint  
        private string eventNo;
        private ArrayList sharePointFolderList;  // Eine Liste aller Ordner der Liste
        private SharePointFile sharePointFile; // Containerklasse um Dateien von SharePoint abzurufen
        private SharePointFolder sharePointFolder;
        #endregion

        #region Constructors
        //Konstruktor
        public SharePointFileGetter() { }

        [Obsolete("Use SharePointFileGetter(string, string, string, string, string, string, string) instead")]
        public SharePointFileGetter(string baseUrl, string subWebsite, string user, string password, string subPage, string listName) {

            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.url = this.baseUrl + subWebsite;
            this.subPage = subPage;
            this.listName = listName;
            this.sharePointFolderList = new ArrayList();
            try {
                if (String.IsNullOrEmpty(subPage)) {
                    Connector connector = Connector.GetConnector(this.url + this.subPage, user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {
                    Connector connector = Connector.GetConnector(this.url + this.subPage + "/", user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            }
            LoadSharePointData();

        }

        public SharePointFileGetter(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {

            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.url = this.baseUrl + subWebsite;
            this.subPage = subPage;
            this.listName = listName;
            this.sharePointFolderList = new ArrayList();
            try {
                if (String.IsNullOrEmpty(subPage)) {

                    Connector connector = Connector.GetConnector(this.url, user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {

                    Connector connector = Connector.GetConnector(this.url + this.subPage + "/", user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }

                this.filter = filter;

                LoadSharePointData();

            } catch (ArgumentNullException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (UriFormatException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (NotSupportedException ex) {
                logger.Error(ex.Message);
                logger.Info("Credentials");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }

        #endregion

        // Lädt die SharePoint-Seite und filtert auf die Liste mit den Dateien.
        // Speichert die Namen der obersten Ordner in die Liste folderNameList.
        // Gibt dann alle Ordner der Liste weiter an LoadSubFolders weiter.
        private void LoadSharePointData() {

            List list = this.site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();

            Folder rootFolder = list.RootFolder;
            this.clientContext.Load(rootFolder);
            this.clientContext.ExecuteQuery();


            FolderCollection allFolders = rootFolder.Folders;
            if (String.IsNullOrEmpty(this.filter)) {

                GetFolderNamesWithoutFilter(allFolders);
                if (this.sharePointFolderList.Count > 0)
                    LoadSubFolders(allFolders);
            } else {

                GetFolderNamesWithFilter(allFolders);
                if (this.sharePointFolderList.Count > 0)
                    LoadSubFoldersWithFilter();
            }

            /****Aufräumarbeit****/
            Connector.Disconnect();
            this.clientContext.Dispose();
            this.site = null;
            return; // Nachdem LoadSubFolder abgeschlossen ist gibt es nichts weiteres zutun.

        }

        #region LoadSubfolders
        // LoadSubFolders lädt die Unterordner der Liste und überprüft ob es 
        // Unterordner in den Unterordnern gibt. Wenn es welche gibt wird die 
        // LoadSubFolders-Funktion rekursive aufgerufen, bis es nur noch Ordner
        // ohne Unterordner gibt. Dann werden die Ordner an die LoadFiles-Funktion
        // weitergegeben.
        private void LoadSubFoldersWithFilter() {
            this.clientContext.Load(this.filteredFolder.Folders);
            this.clientContext.ExecuteQuery();
            if (this.filteredFolder.Folders.Count > 0) {
                foreach (Folder subFolder in this.filteredFolder.Folders) {
                    if (subFolder.Name == "Forms") { //Forms ist von SharePoint
                        continue;
                    }
                    foreach (SharePointFolder folder in this.sharePointFolderList) {
                        this.clientContext.Load(subFolder.ParentFolder);
                        this.clientContext.ExecuteQuery();
                        if (folder.GetFolderName() == subFolder.ParentFolder.Name) {
                            this.sharePointFolder = new SharePointFolder(subFolder.Name, subFolder.ServerRelativeUrl, folder);
                            folder.AddSubFolder(this.sharePointFolder);
                        }
                    }

                    if (CheckForSubFolders(subFolder))
                        LoadSubFolders(subFolder.Folders);

                    LoadFiles(subFolder);
                }
            }
            this.clientContext.Load(this.filteredFolder.Files);
            this.clientContext.ExecuteQuery();
            if (this.filteredFolder.Files.Count > 0) {
                LoadFiles(this.filteredFolder);
            }
        }

        private void LoadSubFolders(FolderCollection allFolders) {

            this.clientContext.Load(allFolders);
            this.clientContext.ExecuteQuery();

            foreach (Folder subFolder in allFolders) {
                if (subFolder.Name == "Forms") {//Forms ist von SharePoint
                    continue;
                }
                foreach (SharePointFolder folder in this.sharePointFolderList) {
                    this.clientContext.Load(subFolder.ParentFolder);
                    this.clientContext.ExecuteQuery();
                    if (folder.GetFolderName() == subFolder.ParentFolder.Name) {
                        this.sharePointFolder = new SharePointFolder(subFolder.Name, subFolder.ServerRelativeUrl, folder);
                        folder.AddSubFolder(this.sharePointFolder);
                    }
                }

                if (CheckForSubFolders(subFolder))
                    LoadSubFolders(subFolder.Folders);

                LoadFiles(subFolder);
            }
        }
        #endregion

        // LoadFiles bekommt einen Ordner übergeben und lädt die 
        // Dateien, die in dem Ordner liegen. Dann wird für jede Datei
        // in dem Ordner eine neue Instanz von der Klasse SharePointFile
        // erstellt und die Instanzvariabelen gesetzt.
        private void LoadFiles(Folder subFolder) {

            
            string subFolderName = subFolder.Name;
            this.clientContext.Load(subFolder.Files);
            this.clientContext.ExecuteQuery();

            foreach (File file in subFolder.Files) {
                // aspx Files sind SharePoint Konfigurationsdateien und für uns unwichtig
                if (file.Name.Contains(".aspx"))
                    continue;

                this.sharePointFile = new SharePointFile();
                this.sharePointFile.SetName(file.Name);
                this.sharePointFile.BuildLinkingUrl(this.baseUrl.Substring(0, this.baseUrl.Length - 1)); // Die BaseUrl hat am Ende ein '/', das zuviel wäre,
                this.sharePointFile.BuildLinkingUrl(file.ServerRelativeUrl);                             // denn ServerRelativeUrl beginnt mit einem '/'.

                foreach (SharePointFolder folder in this.sharePointFolderList) {

                    CheckForMatchingSubFolder(subFolderName, folder);
                }
            }
        }

        // Überprüft ob in dem übergebenen Folder weitere SubFolder sind.
        private bool CheckForSubFolders(Folder subFolder) {

            this.clientContext.Load(subFolder.Folders);
            this.clientContext.ExecuteQuery();

            if (subFolder.Folders.Count > 0) {

                return true;
            } else {

                return false;
            }
        }

        #region GetFolderNames
        // Speichert die Ordnernamen der Obersten Ordner in der SharePoint-Liste.
        private void GetFolderNamesWithoutFilter(FolderCollection allFolders) {

            this.clientContext.Load(allFolders);
            this.clientContext.ExecuteQuery();

            foreach (Folder folder in allFolders) {
                if (folder.Name == "Forms") {
                    continue;
                }

                this.sharePointFolder = new SharePointFolder(folder.Name, folder.ServerRelativeUrl, null);
                this.sharePointFolderList.Add(this.sharePointFolder);
            }
        }

        private void GetFolderNamesWithFilter(FolderCollection allFolders) {
            try {
                Folder rootFolder = this.site.Lists.GetByTitle(this.listName).RootFolder;
                this.clientContext.Load(rootFolder, rf => rf.ServerRelativeUrl);
                this.clientContext.ExecuteQuery();
                Folder folder = this.site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + this.filter);
                this.clientContext.Load(folder);
                this.clientContext.ExecuteQuery();

                this.sharePointFolder = new SharePointFolder(folder.Name, folder.ServerRelativeUrl, null);
                this.sharePointFolderList.Add(this.sharePointFolder);
                this.filteredFolder = folder;
               
            } catch (Exception ex) {
                if (ex is System.IO.FileNotFoundException || ex is ServerException) {
                    this.sharePointFolderList.Clear();
                    this.sharePointFolder = null;
                    this.filteredFolder = null;
                } else {
                    throw ex;
                }
            }
        }
        #endregion

        // Filtert auf den Namen Ordners, wenn kein Filter
        // eingestellt wurde gibt die Funktion die obersten Ordner zurück.
        // Gefilterte Dateien oder Ordner werden zurückgegeben
        public ArrayList GetFiles() {

            if (String.IsNullOrEmpty(this.filter)) {

                return this.sharePointFolderList;
            } else {

                ArrayList filteredFiles = new ArrayList();

                foreach (SharePointFolder folder in this.sharePointFolderList) {

                    if (folder.GetFolderName() == this.filter) {
                        if (folder.GetSubFolders().Count > 0)
                            foreach (SharePointFolder subFolder in folder.GetSubFolders()) {

                                filteredFiles.Add(subFolder);
                            }
                        if (folder.GetFiles().Count > 0)
                            foreach (SharePointFile file in folder.GetFiles()) {

                                filteredFiles.Add(file);
                            }
                        return filteredFiles;
                    } else {
                        if (folder.GetSubFolders().Count > 0)
                            foreach (SharePointFolder subfolder in folder.GetSubFolders()) {

                                filteredFiles.AddRange(CheckForMatchingSubFolder(subfolder));
                            }
                    }
                }

                return filteredFiles;
            }
        }

        public ArrayList GetEventFiles() {

            if (String.IsNullOrEmpty(this.eventNo)) {

                return this.sharePointFolderList;
            } else {

                ArrayList filteredFiles = new ArrayList();

                foreach (SharePointFolder folder in this.sharePointFolderList) {

                    if (folder.GetFolderName() == this.eventNo) {
                        if (folder.GetSubFolders().Count > 0)
                            foreach (SharePointFolder subFolder in folder.GetSubFolders()) {

                                filteredFiles.Add(subFolder);
                            }
                        if (folder.GetFiles().Count > 0)
                            foreach (SharePointFile file in folder.GetFiles()) {

                                filteredFiles.Add(file);
                            }
                        return filteredFiles;
                    } else {
                        if (folder.GetSubFolders().Count > 0)
                            foreach (SharePointFolder subfolder in folder.GetSubFolders()) {

                                filteredFiles.AddRange(CheckForMatchingSubFolder(subfolder));
                            }
                    }
                }

                return filteredFiles;
            }
        }

        #region CheckForMatchingSubfolder
        // Diese Funktion ist eine Überladung  ist eine der CheckForMatchingSubFolder sie nimmt nur einen SharePointFolder an
        // und gibt eine ArrayList zurück. Auch hier wird nach einem passenden Ordnernamen gesucht und wenn er gefunden wird,
        // werden die Unterordner und Dateien, die an dem Ordner hängen zurückgegeben.
        private ArrayList CheckForMatchingSubFolder(SharePointFolder folder) {
            ArrayList filteredFiles = new ArrayList();

            if (folder.GetFolderName() == this.filter) {
                if (folder.GetSubFolders().Count > 0)
                    foreach (SharePointFolder subFolder in folder.GetSubFolders()) {

                        filteredFiles.Add(subFolder);
                    }
                if (folder.GetFiles().Count > 0)
                    foreach (SharePointFile file in folder.GetFiles()) {

                        filteredFiles.Add(file);
                    }
                return filteredFiles;
            } else {
                if (folder.GetSubFolders().Count > 0)
                    foreach (SharePointFolder subFolder in folder.GetSubFolders())
                        CheckForMatchingSubFolder(subFolder);

            }
            return filteredFiles;
        }

        // Die Funktion CheckForMatchingSubFolder sucht die Unterordner des übergebenen Ordners nach
        // dem passenen Namen ab, der auch übergeben wird. Wenn ein Unterordner ebensfalls Unterordner 
        // aufweist wird die Funktion rekursiv aufgerufen. Wenn der passene Ordner gefunden wird, wird
        // die zuletzt erstellte SharePointFile an den Ordner gehangen.
        private void CheckForMatchingSubFolder(string subFolderName, SharePointFolder folder) {
            if (folder.GetSubFolders().Count > 0) {
                foreach (SharePointFolder subFolder in folder.GetSubFolders()) {
                    if (subFolder.GetSubFolders().Count > 0)
                        CheckForMatchingSubFolder(subFolderName, subFolder);
                    if (subFolder.GetFolderName() == subFolderName) {
                        this.sharePointFile.SetParentFolder(subFolder);
                        subFolder.AddFile(this.sharePointFile);
                    }
                }
            } else {
                if (folder.GetFolderName() == subFolderName) {
                    this.sharePointFile.SetParentFolder(folder);
                    folder.AddFile(this.sharePointFile);
                }
            }
        }
        #endregion

        private void GetEventWithFilter(FolderCollection allFolders) {

            this.clientContext.Load(allFolders);
            this.clientContext.ExecuteQuery();

            foreach (Folder folder in allFolders) {
                if (folder.Name == "Forms") {
                    continue;
                }
                if (folder.Name != this.eventNo) {
                    continue;
                }
                this.sharePointFolder = new SharePointFolder(folder.Name, folder.ServerRelativeUrl, null);
                this.sharePointFolderList.Add(this.sharePointFolder);
               
                this.filteredFolder = folder;
                break;
            }
        }


        private void GetIFUEvent() {
            List list = this.site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();
            Folder rootFolder = list.RootFolder;
            this.clientContext.Load(rootFolder);
            this.clientContext.ExecuteQuery();

            Folder eventFolder;
           
            if (String.IsNullOrEmpty(subPage)) {
                
                eventFolder = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + this.filter );
            } else {
                
                eventFolder = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/" + this.filter);
            }

            GetEventWithFilter(eventFolder.Folders);

            this.clientContext.Load(eventFolder);
            this.clientContext.ExecuteQuery();
            
            LoadFiles(this.filteredFolder);
        }

    
    

        public void LoadIFUEvent(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter, string eventNo) {
            this.baseUrl = baseUrl;
            this.url = this.baseUrl + subWebsite;
            this.subPage = subPage;
            this.listName = listName;
            this.sharePointFolderList = new ArrayList();
            try {
                if (subPage.Length == 0) {
                    Connector connector = new Connector(this.url, @user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {
                    Connector connector =  new Connector(this.url + this.subPage + "/", @user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }

                this.filter = filter;
                this.eventNo = eventNo;

                GetIFUEvent();

            } catch (ArgumentNullException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (UriFormatException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (NotSupportedException ex) {
                logger.Error(ex.Message);
                logger.Info("Credentials");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            } finally {
                Connector.Disconnect();
            }
        }

        public void OnlyGetEventFolder(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter, string eventNo) {

            this.baseUrl = baseUrl;
            this.url = this.baseUrl + subWebsite;
            this.subPage = subPage;
            this.listName = listName;
            this.sharePointFolderList = new ArrayList();
            try {
                if (subPage.Length == 0) {
                    Connector connector = Connector.GetConnector(this.url, user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {
                    Connector connector = Connector.GetConnector(this.url + this.subPage + "/", user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }

                this.filter = filter;
                this.eventNo = eventNo;

                LoadEventOnly();

            } catch (ArgumentNullException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (UriFormatException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (NotSupportedException ex) {
                logger.Error(ex.Message);
                logger.Info("Credentials");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }

        public void OnlyGetParentFolder(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {

            this.baseUrl = baseUrl;
            this.url = this.baseUrl + subWebsite;
            this.subPage = subPage;
            this.listName = listName;
            this.sharePointFolderList = new ArrayList();
            try {
                if (subPage.Length == 0) {
                    Connector connector = Connector.GetConnector(this.url, user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                } else {
                    Connector connector = Connector.GetConnector(this.url + this.subPage + "/", user, password);
                    this.clientContext = connector.GetClientContext();
                    this.site = connector.GetWebSite();
                }

                this.filter = filter;

                LoadParentOnly();

            } catch (ArgumentNullException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (UriFormatException ex) {
                logger.Error(ex.Message);
                logger.Info("Clientcontext Uri");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (NotSupportedException ex) {
                logger.Error(ex.Message);
                logger.Info("Credentials");
                logger.Debug("Url: " + this.url + " SubPage: " + this.subPage + " User: " + user);
            } catch (Exception ex) {
                logger.Error(ex.Message);
                logger.Debug(ex.StackTrace);
            }
        }

        // Diese Funktion sucht auf dem SharePoint nach der Liste, 
        // die übergeben wurde und ruft dann die GetFolderNamesWithFilter-
        // Funktion auf um den Parentfolder zu finden und in die sharePointFolderList
        // zu speichern.
        private void LoadParentOnly() {

            List list = this.site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();

            Folder rootFolder = list.RootFolder;
            this.clientContext.Load(rootFolder);
            this.clientContext.ExecuteQuery();


            FolderCollection allFolders = rootFolder.Folders;

            GetFolderNamesWithFilter(allFolders);

            /****Aufräumarbeit****/
            Connector.Disconnect();
            this.clientContext.Dispose();
            this.site = null;
        }

        private void  LoadEventOnly() {
            List list = this.site.Lists.GetByTitle(this.listName);
            this.clientContext.Load(list);
            this.clientContext.ExecuteQuery();

            Folder rootFolder = list.RootFolder;
            this.clientContext.Load(rootFolder);
            this.clientContext.ExecuteQuery();


            Folder eventFolder = site.GetFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl+ "/" + this.filter);
            this.clientContext.Load(eventFolder.Folders);
            this.clientContext.ExecuteQuery();
            GetEventWithFilter(eventFolder.Folders);
        }

        /**********Getter/Setter**********/
        public ArrayList GetSharePointFolderList() { return this.sharePointFolderList; }
        public void SetFilter(string filter) { this.filter = filter; }
        public SharePointFolder GetParentFolder() {
            if (this.sharePointFolderList.Count > 0){
                return this.sharePointFolderList[0] as SharePointFolder;
            } else {
                return null;
            }
        }
    }
}
