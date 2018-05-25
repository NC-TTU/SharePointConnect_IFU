using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{
    [Serializable]
    public class DragDropEventArgs : EventArgs
    {
        public DragDropEventArgs(string filePath, string parentFolder) {
            this.FilePath = filePath;
            this.ParentFolder = parentFolder;
        }

        public string FilePath { get; }
        public string ParentFolder { get; }
    }
}
