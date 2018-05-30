using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{   
    [Serializable]
    public class FeeAttachment
    {
        readonly string link;
        readonly string name;

        public FeeAttachment(string link, string name) {
            this.link = link;
            this.name = link.Split('/')[link.Split('/').Length - 1].Split('.')[0];
        }

        public string GetLink() { return this.link; }
        public string GetName() { return this.name; }
    }
}
