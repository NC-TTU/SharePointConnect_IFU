using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{
    [Serializable]
    public class BrochureOrder
    {
        readonly Guid guid;
        readonly string filePath;
        readonly string fileExtension;
        readonly string fileName;
        readonly string contactNo;
        readonly string eventNo;
        readonly string eventTemplateNo;
        readonly string courseNo;
        readonly string articleNo;
        readonly int status;


        public BrochureOrder(Guid guid, string contactNo, string eventNo, string eventTemplateNo, string courseNo, string articleNo, int status) {
            this.guid = guid;
            this.contactNo = contactNo;
            this.eventNo = eventNo;
            this.eventTemplateNo = eventTemplateNo;
            this.courseNo = courseNo;
            this.articleNo = articleNo;
            this.status = status;
        }


        public BrochureOrder(Dictionary<string, object> brochureOrder, string filePath) {
            this.guid = Guid.Empty;
            this.filePath = filePath;
            this.fileExtension = String.Empty;
            this.fileName = String.Empty;
            this.contactNo = String.Empty;
            this.eventNo = String.Empty;
            this.eventTemplateNo = String.Empty;
            this.courseNo = String.Empty;
            this.articleNo = String.Empty;
            this.status = 0; // Status 0, entspricht 'Neu' in NAV

            foreach (KeyValuePair<string, object> pair in brochureOrder) {
                if (pair.Value != null) {
                    switch (pair.Key) {
                        case "GUID":
                            this.guid = (Guid)pair.Value;
                            break;
                        case "FileLeafRef":
                            this.fileName = pair.Value.ToString();
                            break;
                        case "File_x0020_Type":
                            this.fileExtension = pair.Value.ToString();
                            break;
                        case "IFUKuAnmldgContactNumber":
                            this.contactNo = pair.Value.ToString();
                            break;
                        case "IFUEventnumber":
                            this.eventNo = pair.Value.ToString();
                            break;
                        case "IFUEventTemplateNumber":
                            this.eventTemplateNo = pair.Value.ToString();
                            break;
                        case "IFUCourseNumber":
                            this.courseNo = pair.Value.ToString();
                            break;
                        case "IFUArticleNumber":
                            this.articleNo = pair.Value.ToString();
                            break;
                    }
                }
            }

            this.fileName.Replace("." + this.fileExtension.ToLower(), "");
        }


        /****GETTER****/
        public Guid GetGuid() { return this.guid; }
        public string GetFilePath() { return this.filePath; }
        public string GetFileExtension() { return this.fileExtension; }
        public string GetFileName() { return this.fileName; }
        public string GetContactNo() { return this.contactNo; }
        public string GetEventNo() { return this.eventNo; }
        public string GetEventTemplateNo() { return this.eventTemplateNo; }
        public string GetCourseNo() { return this.courseNo; }
        public string GetArticleNo() { return this.articleNo; }
        public int GetStatus() { return this.status; }
    }
}
