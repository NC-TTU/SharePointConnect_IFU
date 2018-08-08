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

        public Invoice(string vendorNo, string externalDocNo, decimal amount, string barcode) {
            this.vendorNo = vendorNo;
            this.externalDocNo = externalDocNo;
            this.amount = amount;
            this.barcode = barcode;
        }

        public Invoice(Dictionary<string,object> invoice) {
            this.vendorNo = String.Empty;
            this.externalDocNo = String.Empty;
            this.amount = 0.0m;
            this.barcode = String.Empty;

            foreach(KeyValuePair<string,object> pair in invoice) {
                if(pair.Key == "PLACEHOLDERVENDORNO") { // TODO
                    this.vendorNo = pair.Value.ToString();
                }
                if (pair.Key == "IFUInvoiceInvoiceNumber") {
                    this.externalDocNo = pair.Value.ToString();
                }
                if (pair.Key == "IFUInvoiceTotal") {
                    this.amount = (decimal)pair.Value;
                }
                if(pair.Key == "IFUInvoiceBarcode") {
                    this.barcode = pair.Value.ToString();
                }
            }
        }

        public string GetVendorNo() { return this.vendorNo; }
        public string GetExternalDocNo() { return this.externalDocNo; }
        public decimal GetAmount() { return this.amount; }
        public string GetBarcode() { return this.barcode; }

    }
}
