/// <summary>
/// Letzte Änderung am 08.12.2017   
/// --Umstellung zur Nutzung von SharePointFolder-Klasse
///  
/// Erstellt von: NC-TTU
/// Erstellt am: 01.12.2017   
/// 
/// Die SharePointFile-Klasse ist eine einfache Containerklasse, die die
/// wichtigen Daten von den SharePoint-Dokumenten speichern soll.
/// </summary>



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{


    public class SharePointFile
    {

        private string name;
        private string linkingUrl;
        private SharePointFolder parentFolder;

        public SharePointFile() { }

        public SharePointFile(SharePointFolder parentFolder) {

            this.parentFolder = parentFolder;
        }

        public void BuildLinkingUrl(string part) {
            this.linkingUrl += part;
        }

        /***********Getter/Setter************/
        public string GetName() { return this.name; }
        public string GetLinkingUrl() { return this.linkingUrl; }
        public SharePointFolder GetParentFolder() { return this.parentFolder; }

        public void SetName(string name) { this.name = name; }
        public void SetLinkingUrl(string linkingUrl) { this.linkingUrl = linkingUrl; }
        public void SetParentFolder(SharePointFolder parentFolder) { this.parentFolder = parentFolder; }

        new public string ToString() {
            StringBuilder builder = new StringBuilder();

            builder.Append("Filename:" + this.GetName() + "\n");
            builder.Append("LinkingUrl:" + this.GetLinkingUrl() + "\n\n");

            return builder.ToString();
        }

        public string Display { get { return String.Format("{0,20}", this.name); } }
    }
}
