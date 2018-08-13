﻿using System;
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

        public Invoice(string vendorNo, string externalDocNo, decimal amount, string barcode, DateTime dueDate) {
            this.vendorNo = vendorNo;
            this.externalDocNo = externalDocNo;
            this.amount = amount;
            this.barcode = barcode;
            this.dueDate = dueDate;
        }

        public Invoice(Dictionary<string,object> invoice) {
            this.vendorNo = String.Empty;
            this.externalDocNo = String.Empty;
            this.amount = 0.0m;
            this.barcode = String.Empty;
            this.dueDate = default(DateTime);

            foreach(KeyValuePair<string, object> pair in invoice) {
                if (pair.Value != null) {
                    switch (pair.Key) {
                        case "PLACEHOLDERVENDORNO": // TODO
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
                        case "PLACEHOLDERDUEDATE":
                            this.dueDate = DateTime.Parse(pair.Value.ToString());
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

    }
}
