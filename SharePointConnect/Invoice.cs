using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{
    [Serializable]
    public class Invoice
    {
        readonly string vendorNo;
        readonly string externalDocNo;
        readonly decimal amount;
        readonly string barcode;
        readonly DateTime dueDate;
        readonly DateTime createdAt;
        readonly bool creditMemo;
        readonly string eventNo;


        public Invoice(string vendorNo, string externalDocNo, decimal amount, string barcode, DateTime dueDate, DateTime createdAt, bool creditMemo, string eventNo) {
            this.vendorNo = vendorNo;
            this.externalDocNo = externalDocNo;
            this.amount = amount;
            this.barcode = barcode;
            this.dueDate = dueDate;
            this.createdAt = createdAt;
            this.creditMemo = creditMemo;
            this.eventNo = eventNo;
        }

        public Invoice(Dictionary<string,object> invoice) {
            this.vendorNo = String.Empty;
            this.externalDocNo = String.Empty;
            this.amount = 0.0m;
            this.barcode = String.Empty;
            this.dueDate = default(DateTime);
            this.createdAt = default(DateTime);
            this.eventNo = String.Empty;

            foreach(KeyValuePair<string, object> pair in invoice) {
                if (pair.Value != null) {
                    switch (pair.Key) {
                        case "IFUInvoiceSupplierNr": 
                            this.vendorNo = pair.Value.ToString();
                            break;
                        case "IFUInvoiceInvoiceNumber":
                            this.externalDocNo = pair.Value.ToString();
                            break;
                        case "IFUInvoiceTotal":
                            this.amount = Decimal.Parse(pair.Value.ToString());
                            break;
                        case "IFUInvoiceBarcode":
                            this.barcode = pair.Value.ToString();
                            break;
                        case "IFUInvoiceDueDate":
                            this.dueDate = DateTime.Parse(pair.Value.ToString());
                            break;
                        case "Created":
                            this.createdAt = DateTime.Parse(pair.Value.ToString());
                            break;
                        case "Gutschrift":
                            this.creditMemo = (bool)pair.Value;
                            break;
                        case "IFUEventnumber":
                            this.eventNo = pair.Value.ToString();
                            break;
                    }
                }
            }
        }

        public string GetVendorNo() { return this.vendorNo; }
        public string GetExternalDocNo() { return this.externalDocNo; }
        public decimal GetAmount() { return this.amount; }
        public string GetBarcode() { return this.barcode; }
        public DateTime GetDueDate() { return this.dueDate; }
        public DateTime GetCreatedAt() { return this.createdAt; }
        public bool IsCreditMemo() { return this.creditMemo; }
        public string GetEventNo() {return this.eventNo; }
    }
}
