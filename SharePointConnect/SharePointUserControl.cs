///<summary>
///Änderung am 31.01.18
/// -- Implementierung von ConnectToSPButton, welcher nach 
///    nach Klick die Ordnerstruktur aufbaut.
/// 
/// Änderung am 22.12.17                                 
/// -- Implementierung eines Buttons der die Liste auf      
///    SharePoint öffnet. Upload von Files in SharePoint. 
///    
/// Änderung am: 15.12.17                                    
/// -- Darstellung des Ordnerverlaufs als Treeview          
///    (TWI geile Idee!!) 
///    
///    Erstellt von: NC-TTU  
///    Erstellt am: 14.12.17 
///    
///    Die SharePointUserControl-Klasse definiert die Aktionen,
///    die auf dem Windows-Form ausgeführt werden können.
///    Außerdem ist sie dafür verantwortlich den User zu        
///    erlauben Ordner und Dateien auszuwählen und zu öffnen.
/// </summary>


using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Resources;
using log4net;

namespace SharePointConnect
{

    public partial class SharePointUserControl : UserControl
    {

        private static readonly ILog logger = LogManager.GetLogger(typeof(SharePointUserControl));

        #region Instancevaribles
        private System.Collections.ArrayList sharePointDataList;  // Den ganzen gefilterten Ordnerverlauf von SharePoint
        private SharePointFolder parent;
        private string baseUrl;
        private string subWebsite;
        private string user;
        private string password;
        private string subPage;
        private string listName;
        private string filter;
        private string eventNo;
        List<TreeNode> folderList;
        List<TreeNode> fileList;
        private static ResourceManager resourceManager;

        #endregion

        public event MethodInvoker Create = delegate { };
        public event EventHandler<DragDropEventArgs> Drag_Drop = delegate { };

        public SharePointUserControl() {

            // CultureInfo fragt nach der momentanen Sprache vom System.
            // Wenn die Sprache Deutsch ist wird die lokale deutsche resx-Datei verwendet,
            // bei allen anderen Sprachen wird die lokale englische resx-Datei verwendet.
            // In den resx-Dateien sind ein paar Übersetzungen für Fehlerausgaben hinterlegt.
            CultureInfo currentCulture = CultureInfo.CurrentCulture;
            if (currentCulture.Name.Equals("de-DE")) {
                resourceManager = new ResourceManager("SharePointConnect.de_local", Assembly.GetExecutingAssembly());
            } else {
                resourceManager = new ResourceManager("SharePointConnect.en_local", Assembly.GetExecutingAssembly());
            }

            Assembly assembly;
            Stream imageStream;
            this.folderList = new List<TreeNode>();
            this.fileList = new List<TreeNode>();
            InitializeComponent();
            TWI_TreeView.Visible = false;
            OpenSPButton.Visible = false;
            InitSharePointDataList();
            TWI_TreeView.AllowDrop = true;
            TWI_TreeView.NodeMouseDoubleClick += SharePointTreeView_DoubleClick;
            TWI_TreeView.DragEnter += TWI_TreeView_DragEnter;
            TWI_TreeView.DragDrop += TWI_TreeView_DragDrop;
            OpenSPButton.Text = resourceManager.GetString("toSharePoint");
            ConnectButton.Text = resourceManager.GetString("connect");

            // Hier werden die beiden Images, die im Assembly hinterlegt sind geladen und der Imagelist des Treeviews zugefügt. 
            try {

                assembly = Assembly.GetExecutingAssembly();

                ImageList imageList = new ImageList {
                    ImageSize = new Size(20, 20)
                };
                imageList.Images.Add(Image.FromStream(imageStream = assembly.GetManifestResourceStream("SharePointConnect.Images.folder.png")));
                imageList.Images.Add(Image.FromStream(imageStream = assembly.GetManifestResourceStream("SharePointConnect.Images.file.png")));

                TWI_TreeView.ImageList = imageList;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
                logger.Error(ex.Message + "\n" + ex.StackTrace);
            }
        }

        #region TreeViewFunktionen
        // Eventhandler für das betreten des WindowsForm mit Dateien.
        private void TWI_TreeView_DragEnter(object sender, DragEventArgs e) {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        // Eventhandler für den Drop von Dateien in das WindowsForm bzw. den Treeview
        // hier wird der Dateipfad übergeben, über den wir an die Dateien kommen, die hochgeladen werden sollen.
        // Die Dateipfade werden in das Stringarray "files" gespeichert und in der foreach-Schleife durchlaufen.
        // Am Anfang der Funktion wird die Klasse SharePointFileUploader initialisiert, damit die Klasse die ganze
        // Funktion über ansprechbar ist.
        private void TWI_TreeView_DragDrop(object sender, DragEventArgs e) {

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);


            //SharePointFileUploader fileUploader = new SharePointFileUploader(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter);

            foreach (string file in files) {

                SharePointFile newSPFile = new SharePointFile();

                string[] parts = file.Split('\\'); // Teilt die Dateipfade am Trennstrich '\' und speichert diese in das stringarray parts

                newSPFile.SetName(parts[parts.Length - 1]); // Das letzt Element des Arrays ist der Name der Datei.


                Point targetPoint = TWI_TreeView.PointToClient(new Point(e.X, e.Y)); // Hier wird bestimmt wo die Dateien im Treeview gedroped wurden.
                TreeNode parentNode = TWI_TreeView.GetNodeAt(targetPoint);

                if (parentNode != null) {

                    // Hier wird geprüft ob die gedropte Datei auf einem TreeNode gedropt wurde,
                    // der in der fileList drin ist. Wenn dem so ist wird der parentNode auf den Parent
                    // von sich selbst verwiesen, weil jede Datei in einem Ordner liegen muss.
                    foreach (TreeNode node in this.fileList)
                        if (node.Text == parentNode.Text)
                            parentNode = parentNode.Parent;
                }

                SharePointFolder parentFolder;

                try {
                    parentFolder = FindParentFolder(parentNode.Text);
                } catch (ArgumentException ex) {
                    MessageBox.Show(ex.Message);
                    logger.Error(ex.Message + "\n" + ex.StackTrace);
                    logger.Debug("ParentNodeText: " + parentNode.Text);
                    return;
                }
                Drag_Drop(this, new DragDropEventArgs(file, parentFolder.GetServerRelativeUrl()));
                /*fileUploader.SetUrl(parentFolder);

                // Hier wird die Datei über die Klasse SharePointFileUploader hochgeladen.
                try {
                    fileUploader.SetConnection();
                    fileUploader.UploadFile(file, parentFolder);

                } catch (Exception ex) {
                    MessageBox.Show(resourceManager.GetString("errorUploading") + newSPFile.GetName() + "\n" + resourceManager.GetString("filepath") + file + "\n\n" + ex.Message);
                    logger.Error(ex.Message + "\n" + ex.StackTrace);
                    logger.Debug(resourceManager.GetString("errorUploading") + newSPFile.GetName() + "\n" + resourceManager.GetString("filepath") + file + "\n\n" + ex.Message);
                    return;
                    }
            }
            // Am Ende wird die aktualisierte Liste neu geladen.
            this.AddSharePointData(fileUploader.UpdateSharePointDataList());
                */
            }
        }

        // In der Funktion FillTWI_TreeView wird der Treeview aufgebaut es wird der oberste Ordner hinzugefügt 
        // und geprüft ob dieser noch Unterordner besitzt. Wenn ja werden die Unterordner ebenfalls auf Unterordner und Dateien
        // geprüft, falls die Unterordner welche besitzen werden sie in eine Liste gespeichert. Anschließend werden 
        // die Unterordner des Parentordners in den TreeView gespeichert.
        private void FillTWI_TreeView() {
            TWI_TreeView.BeginUpdate();
            TWI_TreeView.Nodes.Add(this.parent.GetFolderName());
            TWI_TreeView.EndUpdate();

            List<SharePointFolder> folderList = new List<SharePointFolder>();

            if (this.parent.GetSubFolders().Count > 0) {
                foreach (SharePointFolder subFolder in this.parent.GetSubFolders()) {
                    if (subFolder.GetSubFolders().Count > 0 || subFolder.GetFiles().Count > 0)
                        folderList.Add(subFolder);

                    TWI_TreeView.BeginUpdate();
                    TWI_TreeView.Nodes[0].Nodes.Add(subFolder.GetFolderName(), subFolder.GetFolderName(), 0, 0);
                    this.folderList.Add(new TreeNode(subFolder.GetFolderName()));
                    TWI_TreeView.EndUpdate();

                }
            }

            if (this.parent.GetFiles().Count > 0) {
                foreach (SharePointFile file in this.parent.GetFiles()) {

                    TWI_TreeView.BeginUpdate();
                    TWI_TreeView.Nodes[0].Nodes.Add(file.GetName(), file.GetName(), 1, 1);
                    this.fileList.Add(new TreeNode(file.GetName()));
                    TWI_TreeView.EndUpdate();
                }
            }

            if (folderList.Count > 0) { // folderList ist die Liste mit den Unterordnern des Elternordners, die entweder andere Unterordner oder Dateien enthalten. 

                FillTWI_TreeView(folderList);
            }
        }

        // Hier wird das ganze was wir mit dem Elternordner gemacht haben für die Unterordner und deren 
        // Unterordner bzw. Dateien noch mal getan. Gibt es weitere Unterordner wird die Funktion rekursiv aufgerufen.
        private void FillTWI_TreeView(IList<SharePointFolder> folderList) {

            List<SharePointFolder> subfolderList = new List<SharePointFolder>();

            foreach (SharePointFolder folder in folderList) {
                if (folder.GetSubFolders().Count > 0)
                    subfolderList.Add(folder);

                foreach (SharePointFolder subFolder in folder.GetSubFolders()) {
                    TreeNode parentNode;
                    try {
                        parentNode = FindTreeNode(subFolder.GetFolderName());
                    } catch (Exception ex) {
                        MessageBox.Show(ex.Message);
                        logger.Error(ex.Message + "/n" + ex.StackTrace);
                        logger.Debug("Subfoldername: " + subFolder.GetFolderName());
                        return;
                    }
                    if (parentNode != null) {
                        TWI_TreeView.BeginUpdate();
                        parentNode.Nodes.Add(folder.GetFolderName(), folder.GetFolderName(), 0, 0);
                        this.folderList.Add(new TreeNode(folder.GetFolderName()));
                        TWI_TreeView.EndUpdate();
                    }
                }


                if (folder.GetFiles().Count > 0) {
                    foreach (SharePointFile file in folder.GetFiles()) {
                        TreeNode parentNode;
                        try {
                            parentNode = FindTreeNode(folder.GetFolderName());
                        } catch (Exception ex) {
                            MessageBox.Show(ex.Message);
                            logger.Error(ex.Message + "/n" + ex.StackTrace);
                            logger.Debug("Foldername: " + folder.GetFolderName());
                            return;
                        }
                        if (parentNode != null) {
                            TWI_TreeView.BeginUpdate();
                            parentNode.Nodes.Add(file.GetName(), file.GetName(), 1, 1);
                            this.fileList.Add(new TreeNode(file.GetName()));
                            TWI_TreeView.EndUpdate();
                        }
                    }
                }
            }
            if (subfolderList.Count > 0) {
                FillTWI_TreeView(subfolderList);
            }
        }

        // Diese Methode reagiert auf den Doppelklick des Users, wenn der User
        // einen Doppelklick auf ein Element des TWI_TreeView macht wird der Name
        // übergeben. Dann werden alle Dateien, die in den Files-ArrayLists der
        // SharePointFolder vorhanden sind durchlaufen und es wird überprüft ob
        // der Name übereinstimmt. Wenn der Name übereinstimmt wird die Datei geöffnet.
        private void SharePointTreeView_DoubleClick(object sender, EventArgs e) {

            if (this.folderList.Contains(TWI_TreeView.SelectedNode) || TWI_TreeView.SelectedNode == null) {
                return;
            }

            string selectedItemName = TWI_TreeView.SelectedNode.Name;

            foreach (object obj in sharePointDataList) {
                if (obj.GetType() == typeof(SharePointFolder)) {

                    var folder = obj as SharePointFolder;
                    SharePointFile file;

                    if (folder.GetSubFolders().Count > 0) {
                        try {
                            file = SearchForFileName(selectedItemName, folder.GetSubFolders());
                        } catch (ArgumentException ex) {
                            MessageBox.Show(ex.Message);
                            logger.Error(ex.Message + "\n" + ex.StackTrace);
                            logger.Debug("SelectedItemName: " + selectedItemName + "Foldercount: " + folder.GetSubFolders().Count);
                            foreach (SharePointFolder f in folder.GetSubFolders())
                                logger.Debug("FolderName: " + f.GetFolderName());
                            return;
                        }

                        if (file != null) {
                            try {
                                System.Diagnostics.Process.Start(file.GetLinkingUrl());
                            } catch (FileNotFoundException ex) {
                                MessageBox.Show(resourceManager.GetString("fileNotFound") + file.GetName() + "/n" + ex.Message);
                                logger.Error(ex.Message + "\n" + ex.StackTrace);
                                logger.Debug("FileName: " + file.GetName() + "LinkingUrl: " + file.GetLinkingUrl() + "Parentfolder: " + file.GetParentFolder().GetFolderName());
                            } catch (Exception ex) {
                                MessageBox.Show(resourceManager.GetString("triedToOpen") + file.GetName() + "/n" + ex.Message);
                                logger.Error(ex.Message + "\n" + ex.StackTrace);
                                logger.Debug("FileName: " + file.GetName() + "LinkingUrl: " + file.GetLinkingUrl() + "Parentfolder: " + file.GetParentFolder().GetFolderName());
                            }
                        }
                    }

                    if (folder.GetFiles().Count > 0) {
                        foreach (SharePointFile f in folder.GetFiles()) {

                            if (f.GetName() == selectedItemName) {
                                try {
                                    System.Diagnostics.Process.Start(f.GetLinkingUrl());
                                } catch (FileNotFoundException ex) {
                                    MessageBox.Show(resourceManager.GetString("fileNotFound") + f.GetName() + "/n" + ex.Message + "/n" + ex.StackTrace);
                                    logger.Error(ex.Message + "\n" + ex.StackTrace);
                                    logger.Debug("FileName: " + f.GetName() + "LinkingUrl: " + f.GetLinkingUrl() + "Parentfolder: " + f.GetParentFolder().GetFolderName());
                                } catch (Exception ex) {
                                    MessageBox.Show(resourceManager.GetString("triedToOpen") + f.GetName() + "/n" + ex.Message + "/n" + ex.StackTrace);
                                    logger.Error(ex.Message + "\n" + ex.StackTrace);
                                    logger.Debug("FileName: " + f.GetName() + "LinkingUrl: " + f.GetLinkingUrl() + "Parentfolder: " + f.GetParentFolder().GetFolderName());
                                }
                            }
                        }
                    }
                } else if (obj.GetType() == typeof(SharePointFile)) {

                    var file = obj as SharePointFile;

                    if (file.GetName() == selectedItemName) {
                        try {
                            System.Diagnostics.Process.Start(file.GetLinkingUrl());
                        } catch (FileNotFoundException ex) {
                            MessageBox.Show(resourceManager.GetString("fileNotFound") + file.GetName() + "/n" + ex.Message);
                            logger.Error(ex.Message + "\n" + ex.StackTrace);
                            logger.Debug("FileName: " + file.GetName() + "LinkingUrl: " + file.GetLinkingUrl() + "Parentfolder: " + file.GetParentFolder().GetFolderName());
                        } catch (Exception ex) {
                            MessageBox.Show(resourceManager.GetString("triedToOpen") + file.GetName() + "/n" + ex.Message);
                            logger.Error(ex.Message + "\n" + ex.StackTrace);
                            logger.Debug("FileName: " + file.GetName() + "LinkingUrl: " + file.GetLinkingUrl() + "Parentfolder: " + file.GetParentFolder().GetFolderName());
                        }
                    }
                }
            }
        }

        #region FindTreeNode
        // FindTreeNode bekommt einen Schlüssel mit dem der 
        // Elternknoten gefunden werden kann. Gesucht wird
        // der Elternknoten über LINQ indem wir die erste Ebene
        // des TWI_TreeViews aufrufen, die TreeNodeCollection durchlaufen
        // und überprüfen ob der Text des Knotens gleich unserem Schlüssel ist.
        // wenn er gefunden wird, wird er in das treeNodes-Array gespeichert.
        // Wenn er nicht gefunden wird, wird die nächste Treeview-Ebene an 
        // FindTreeNode(string key, TreeNodeCollection nodeColl) weitergegeben.
        // Wenn kein TreeNode gefunden wird, wird eine Argumentexception geworfen.
        private TreeNode FindTreeNode(string key) {

            TreeNode[] treeNodes = TWI_TreeView.Nodes[0].Nodes
            .Cast<TreeNode>()
            .Where(r => r.Text == key)
            .ToArray();
            if (treeNodes.Count() > 0)
                return treeNodes[0];
            else {
                try {
                    return FindTreeNode(key, TWI_TreeView.Nodes[0].Nodes);
                } catch (ArgumentException ex) {
                    throw ex;
                }
            }
        }

        private TreeNode FindTreeNode(string key, TreeNodeCollection nodeColl) {

            TreeNode[] treeNodes = nodeColl
            .Cast<TreeNode>()
            .Where(tn => tn.Text == key)
            .ToArray();
            if (treeNodes.Count() > 0)
                return treeNodes[0];
            else {
                foreach (TreeNode node in nodeColl)
                    if (node.Nodes.Count > 0)
                        return FindTreeNode(key, node.Nodes);
            }
            throw new ArgumentException(resourceManager.GetString("noTreeNode" + key));
        }
        #endregion
        #endregion

        private void InitSharePointDataList() {

            this.sharePointDataList = new System.Collections.ArrayList();
        }

        // InitSharePointDataList wird nur einmal bei der Übergabe des Ordnerverlaufs von NAV zum Assembly
        // gestartet. Hier werden die BindingLists für Ordner und Dateien zum ersten mal initialisiert außerdem
        // wird die SharePointDataList befüllt.
        private void InitSharePointDataList(System.Collections.ArrayList sharePointData) {

            this.sharePointDataList = new System.Collections.ArrayList(sharePointData);

            FillTWI_TreeView();
        }




        // Einstieg von NAV zum Assembly, der Ordnerverlauf wird übergeben und der Parentordner wird gesetzt.
        public void AddSharePointData(System.Collections.ArrayList sharePointData) {
            SharePointFileGetter fileGetter = new SharePointFileGetter();
            SharePointFolder folder = null;
            if (String.IsNullOrEmpty(eventNo)){
                fileGetter.OnlyGetParentFolder(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter);
                folder = fileGetter.GetParentFolder();
            } else {
                fileGetter.OnlyGetEventFolder(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter,this.eventNo);
                folder = fileGetter.GetParentFolder();
            }
            TWI_TreeView.Nodes.Clear();
            this.parent = null;

            if (sharePointData == null) { return; } 

            if (sharePointData.Count > 0) {

                if (sharePointData[0].GetType() == typeof(SharePointFolder)) {
                    if (((SharePointFolder)sharePointData[0]).GetParentFolder() != null)
                        this.parent = ((SharePointFolder)sharePointData[0]).GetParentFolder();

                } else if (sharePointData[0].GetType() == typeof(SharePointFile)) {
                    if (((SharePointFile)sharePointData[0]).GetParentFolder() != null)
                        this.parent = ((SharePointFile)sharePointData[0]).GetParentFolder();
                }

                ConnectButton.Visible = false;
                TWI_TreeView.Visible = true;
                OpenSPButton.Visible = true;

                InitSharePointDataList(sharePointData);
            } else if(folder == null){
                DialogResult result = MessageBox.Show(resourceManager.GetString("noFiles").Replace("{0}", this.filter), "", MessageBoxButtons.YesNo);
                if (result == DialogResult.No) {
                    TWI_TreeView.Visible = false;
                    OpenSPButton.Visible = false;
                    ConnectButton.Visible = true;
                } else if (result == DialogResult.Yes) {
                    Create();
                }
            } else {
                this.parent = folder;

                ConnectButton.Visible = false;
                TWI_TreeView.Visible = true;
                OpenSPButton.Visible = true;

                InitSharePointDataList(new System.Collections.ArrayList());
            }
        }



        // Sucht die übergebene Ordnerliste nach dem übergebenen Namen ab,
        // wenn der Ordner, der übergeben wurde Unterordner besitzt ruft er
        // die Funktion erneut auf. Wenn die Datei mit dem passenden Namen gefunden
        // wurde wird sie zurückgegeben, falls keine gefunden wurde wird eine ArgumentExeption geworfen.
        private SharePointFile SearchForFileName(string fileName, System.Collections.ArrayList folderList) {
            foreach (object obj in folderList) {
                var folder = obj as SharePointFolder;

                if (folder.GetSubFolders().Count > 0)
                    return SearchForFileName(fileName, folder.GetSubFolders());

                if (folder.GetFiles().Count > 0)
                    foreach (SharePointFile file in folder.GetFiles())
                        if (file.GetName() == fileName)
                            return file;
            }
            throw new ArgumentException(resourceManager.GetString("fileNotFound") + fileName);
        }

        #region ButtonClickEvents

        private void OpenSPButton_Click_1(object sender, EventArgs e) {
            string url = this.baseUrl.Substring(0, this.baseUrl.Length - 1);
            url += this.parent.GetServerRelativeUrl();

            try {
                System.Diagnostics.Process.Start(url);
            } catch (Exception ex) {
                MessageBox.Show(resourceManager.GetString("triedToOpen") + url + "/n" + ex.Message);
                logger.Error(ex.Message + "\n" + ex.StackTrace);
                logger.Debug("Url: " + url);
            }
        }

        private void ConnectButton_Click(object sender, EventArgs e) {
            if (String.IsNullOrEmpty(this.eventNo)) {
                SharePointFileGetter fileGetter = new SharePointFileGetter(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter);
                this.AddSharePointData(fileGetter.GetFiles());
            } else {
                SharePointFileGetter fileGetter = new SharePointFileGetter();
                fileGetter.LoadIFUEvent(this.baseUrl, this.subWebsite, this.user, this.password, this.subPage, this.listName, this.filter, this.eventNo);
                this.AddSharePointData(fileGetter.GetEventFiles());
            }
        }
        #endregion

        // Hier werden die Informationen aus NAV durchgereicht und gespeichert, damit die SharePointFileUploader-Klasse bei einem Drag&Drop-Event
        // die nötigen Informationen hat um eine Verbindung zum SharePoint aufzubauen.
        public void GetConnection(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string filter) {
            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.user = user;
            this.password = password;
            this.subPage = subPage;
            this.listName = listName;
            this.filter = filter;
            this.eventNo = null;
            TWI_TreeView.Visible = false;
            OpenSPButton.Visible = false;
            ConnectButton.Visible = true;
        }

        public void GetConnectionEvent(string baseUrl, string subWebsite, string user, string password, string subPage, string listName, string templateNo, string eventNo) {
            this.baseUrl = baseUrl;
            this.subWebsite = subWebsite;
            this.user = user;
            this.password = password;
            this.subPage = subPage;
            this.listName = listName;
            this.filter = templateNo;
            this.eventNo = eventNo;
            TWI_TreeView.Visible = false;
            OpenSPButton.Visible = false;
            ConnectButton.Visible = true;

        }
            public void RefreshSPC() {
            this.ConnectButton_Click(this, null);
        }

        #region FindParent
        // Hier wird versucht den Elternordner eines Treenodes zu finden.
        // Wenn kein Elternordner gefunden wird, wird eine ArgumentException geworfen.
        private SharePointFolder FindParentFolder(string folderName) {
            if (folderName == this.parent.GetFolderName())
                return this.parent; // Im einfachsten Fall ist es der oberste Ordner in der Liste.

            // Ansonsten müssen wir nach dem Ordner suchen, wir gehen auf den obersten Ordner.
            // Wenn die Unterordner ebenfalls Unterordner haben wird die Überlagerte Funktion
            // FindParentFolder mit dem gesuchten Ordnername und dem Unterordner aufgerufen. 
            foreach (SharePointFolder subFolder in this.parent.GetSubFolders()) {
                if (subFolder.GetSubFolders().Count > 0) {
                    return FindParentFolder(folderName, subFolder);
                } else {
                    if (subFolder.GetFolderName() == folderName)
                        return subFolder;
                }
            }
            throw new ArgumentException(resourceManager.GetString("noParentfolder") + folderName);
        }

        private SharePointFolder FindParentFolder(string folderName, SharePointFolder folder) {
            foreach (SharePointFolder subFolder in folder.GetSubFolders()) {
                if (subFolder.GetSubFolders().Count > 0) {
                    return FindParentFolder(folderName, subFolder);
                } else {
                    if (subFolder.GetFolderName() == folderName)
                        return subFolder;
                }
            }
            throw new ArgumentException(resourceManager.GetString("noParentfolder") + folderName);
        }
        #endregion


        private void SharePointFileListBox_SelectedIndexChanged(object sender, EventArgs e) { }
        private void SharePointFolderListBox_SelectedIndexChanged(object sender, EventArgs e) { }
        private void SharePointUserControl_Load(object sender, EventArgs e) { }

    }
}
