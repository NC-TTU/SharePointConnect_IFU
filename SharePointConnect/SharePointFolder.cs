///<summary>
/// Erstellt von: NC-TTU
/// Erstellt am: 08.12.2017 
/// 
/// Die SharePointFolder-Klasse ist eine einfache Containerklasse, die den 
/// Pfad auf SharePoint wiederspiegeln soll und SharePointFiles bzw. SharePointFolder speichert.   
/// </summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{


    public class SharePointFolder
    {

        private string folderName;
        private string serverRelativeUrl;
        private SharePointFolder parentFolder;
        private System.Collections.ArrayList subFolders;
        private System.Collections.ArrayList files;

        /*******************Konstruktoren*******************/
        public SharePointFolder() {

            this.subFolders = new System.Collections.ArrayList();
            this.files = new System.Collections.ArrayList();
        }

        public SharePointFolder(SharePointFolder parentFolder) {

            this.parentFolder = parentFolder;
            this.subFolders = new System.Collections.ArrayList();
            this.files = new System.Collections.ArrayList();
        }

        public SharePointFolder(string folderName, string serverRelativeUrl, SharePointFolder parentFolder) {

            this.folderName = folderName;
            this.serverRelativeUrl = serverRelativeUrl;
            this.parentFolder = parentFolder;
            this.subFolders = new System.Collections.ArrayList();
            this.files = new System.Collections.ArrayList();
        }


        public void AddFile(SharePointFile file) {
            this.files.Add(file);
        }

        public void AddSubFolder(SharePointFolder folder) {
            this.subFolders.Add(folder);
        }


        /***********************Getter/Setter**********************/

        public string GetFolderName() { return this.folderName; }
        public string GetServerRelativeUrl() { return this.serverRelativeUrl; }
        public SharePointFolder GetParentFolder() { return this.parentFolder; }
        public System.Collections.ArrayList GetSubFolders() { return this.subFolders; }
        public System.Collections.ArrayList GetFiles() { return this.files; }

        public void SetFolderName(string folderName) { this.folderName = folderName; }
        public void SetServerRelativeUrl(string serverRelativeUrl) { this.serverRelativeUrl = serverRelativeUrl; }
        public void SetParentFolder(SharePointFolder parentFolder) { this.parentFolder = parentFolder; }
        public void SetSubFolders(System.Collections.ArrayList subFolders) { this.subFolders = subFolders; }
        public void SetFiles(System.Collections.ArrayList files) { this.files = files; }

        public string Display { get { return String.Format("{0,20}", this.folderName); } }
    }
}
