using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{
    [Serializable]
    public class ScannedDocument
    {
        readonly Guid guid;
        readonly string fileName;
        readonly string filePath;
        readonly string documentType;
        readonly string contactNo;
        readonly string keyword;
        readonly string eventNo;
        readonly string fileExtension;
        readonly string responsibleUser;

        public ScannedDocument(Dictionary<string,object> scannedDocument, string filePath) {
            this.guid = Guid.Empty;
            this.fileName = String.Empty;
            this.filePath = filePath;
            this.documentType = String.Empty;
            this.contactNo = String.Empty;
            this.keyword = String.Empty;
            this.eventNo = String.Empty;
            this.fileExtension = String.Empty;
            this.responsibleUser = String.Empty;

            foreach (KeyValuePair<string,object> pair in scannedDocument) {
                if(pair.Value != null) {
                    switch (pair.Key) {
                        case "GUID":
                            this.guid = (Guid)pair.Value;
                            break;
                        case "FileLeafRef":
                            this.fileName = pair.Value.ToString();
                            break;
                        case "IFUScan2SPDocType":
                            this.documentType = pair.Value.ToString();
                            break;
                        case "IFUScan2SPContactNum":
                            this.contactNo = pair.Value.ToString();
                            break;
                        case "IFUScan2SPKeyword":
                            this.keyword = pair.Value.ToString();
                            break;
                        case "IFUScan2SPEventNum":
                            this.eventNo = pair.Value.ToString();
                            break;
                        case "File_x0020_Type":
                            this.fileExtension = pair.Value.ToString();
                            break;
                        case "IFUZustaendigePerson":
                            Microsoft.SharePoint.Client.FieldUserValue fuv = (Microsoft.SharePoint.Client.FieldUserValue)pair.Value;
                            this.responsibleUser = fuv.LookupValue.ToString();
                            break;

                    }
                }
            }

            this.fileName = this.fileName.Replace("." + this.fileExtension.ToLower(), "");
        }

        public Guid GetGuid() { return this.guid; }
        public string GetFileName() { return this.fileName; }
        public string GetFilePath() { return this.filePath; }
        public string GetDocumentType() { return this.documentType; }
        public string GetContactNo() { return this.contactNo; }
        public string GetKeyword() { return this.keyword; }
        public string GetEventNo() { return this.eventNo; }
        public string GetFileExtension() { return this.fileExtension; }
        public string GetResponsibleUser() { return this.responsibleUser; }
    }
}
