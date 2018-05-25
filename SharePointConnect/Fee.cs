using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointConnect
{
    [Serializable]
    public class Fee
    {
        readonly Guid guid;
        readonly string referentNo;
        readonly string vatId;
        readonly string iban;
        readonly string eventNo;
        readonly string eventDescription;
        readonly DateTime dateFrom;
        readonly DateTime dateTo;
        readonly string host;
        readonly string hotel;
        readonly double fee;
        readonly double days;
        readonly double dailyRate;
        readonly double dailyAllowance;
        readonly double kilometres;
        readonly double carCosts;
        readonly double byOwnCar;
        readonly double byTrain;
        readonly double byPlane;
        readonly double tripAdditionalCosts;
        readonly string other1Description;
        readonly double other1Value;
        readonly string other2Description;
        readonly double other2Value;
        readonly double netTotal;
        readonly double vat;
        readonly double grossTotal;
        readonly int status;
        readonly bool attachment;

        public Fee(Dictionary<string, object> fee, bool hasAttachment) {
            this.guid = Guid.Empty;
            this.referentNo = String.Empty;
            this.vatId = String.Empty;
            this.iban = String.Empty;
            this.eventNo = String.Empty;
            this.eventDescription = String.Empty;
            this.dateFrom = default(DateTime);
            this.dateTo = default(DateTime);
            this.host = String.Empty;
            this.hotel = String.Empty;
            this.fee = 0.0;
            this.days = 0.0;
            this.dailyRate = 0.0;
            this.dailyAllowance = 0.0;
            this.kilometres = 0.0;
            this.carCosts = 0.0;
            this.byOwnCar = 0.0;
            this.byTrain = 0.0;
            this.byPlane = 0.0;
            this.tripAdditionalCosts = 0.0;
            this.other1Description = String.Empty;
            this.other1Value = 0.0;
            this.other2Description = String.Empty;
            this.other2Value = 0.0;
            this.netTotal = 0.0;
            this.vat = 0.0;
            this.grossTotal = 0.0;
            this.status = 1; // Status 1, entspricht 'in Prüfung' in NAV
            this.attachment = hasAttachment;

            foreach (KeyValuePair<string, object> pair in fee) {
                if (pair.Value != null) {
                    switch (pair.Key) {
                        case "GUID":
                            this.guid = (Guid)pair.Value;
                            break;
                        case "Referentnumber":
                            this.referentNo = pair.Value.ToString();
                            break;
                        case "TaxId":
                            this.vatId = pair.Value.ToString();
                            break;
                        case "Iban":
                            this.iban = pair.Value.ToString();
                            break;
                        case "Eventnumber":
                            this.eventNo = pair.Value.ToString();
                            break;
                        case "ReferentFormDesciption":
                            this.eventDescription = pair.Value.ToString();
                            break;
                        case "DateFrom":
                            this.dateFrom = DateTime.Parse(pair.Value.ToString());
                            break;
                        case "DateTo":
                            this.dateTo = DateTime.Parse(pair.Value.ToString());
                            break;
                        case "Organizer":
                            this.host = pair.Value.ToString();
                            break;
                        case "Hotel":
                            this.hotel = pair.Value.ToString();
                            break;
                        case "Fee":
                            this.fee = Double.Parse(pair.Value.ToString());
                            break;
                        case "DailyBenefit":
                            this.dailyAllowance = Double.Parse(pair.Value.ToString());
                            break;
                        case "Days":
                            this.days = Double.Parse(pair.Value.ToString());
                            break;
                        case "DailyRate":
                            this.dailyRate = Double.Parse(pair.Value.ToString());                       
                            break;
                        case "ByOwnCar":
                            this.byOwnCar = Double.Parse(pair.Value.ToString());
                            break;
                        case "Kilometer":
                            this.kilometres = Double.Parse(pair.Value.ToString());
                            break;
                        case "CostCarKm":
                            this.carCosts = Double.Parse(pair.Value.ToString());
                            break;
                        case "ByTrain":
                            this.byTrain = Double.Parse(pair.Value.ToString());
                            break;
                        case "ByPlane":
                            this.byPlane = Double.Parse(pair.Value.ToString());
                            break;
                        case "AdditionalTourCosts":
                            this.tripAdditionalCosts = Double.Parse(pair.Value.ToString());
                            break;
                        case "OtherADescription":
                            this.other1Description = pair.Value.ToString();
                            break;
                        case "OtherAValue":
                            this.other1Value = Double.Parse(pair.Value.ToString());
                            break;
                        case "OtherBDescription":
                            this.other2Description = pair.Value.ToString();
                            break;
                        case "OtherBValue":
                            this.other2Value = Double.Parse(pair.Value.ToString());
                            break;
                        case "AfterTaxValue":
                            this.netTotal = Double.Parse(pair.Value.ToString());
                            break;
                        case "TaxValue":
                            this.vat = Double.Parse(pair.Value.ToString());
                            break;
                        case "PreTax":
                            this.grossTotal = Double.Parse(pair.Value.ToString());
                            break;
                    }
                }
            }

        }


        /***Getter***/
        public Guid GetGuid() { return this.guid; }
        public string GetReferentNo() { return this.referentNo; }
        public string GetVatId() { return this.vatId; }
        public string GetIban() { return this.iban; }
        public string GetEventNo() { return this.eventNo; }
        public string GetEventDescription() { return this.eventDescription; }
        public DateTime GetDateFrom() { return this.dateFrom; }
        public DateTime GetDateTo() { return this.dateTo; }
        public string GetHost() { return this.host; }
        public string GetHotel() { return this.hotel; }
        public double GetFee() { return this.fee; }
        public double GetDays() { return this.days; }
        public double GetDailyRate() { return this.dailyRate; }
        public double GetDailyAllowance() { return this.dailyAllowance; }
        public double GetKilometres() { return this.kilometres; }
        public double GetCarCosts() { return this.carCosts; }
        public double GetByOwnCar() { return this.byOwnCar; }
        public double GetByTrain() { return this.byTrain; }
        public double GetByPlane() { return this.byPlane; }
        public double GetTripAdditionalCosts() { return this.tripAdditionalCosts; }
        public string GetOther1Description() { return this.other1Description; }
        public double GetOther1Value() { return this.other1Value; }
        public string GetOther2Description() { return this.other2Description; }
        public double GetOther2Value() { return this.other2Value; }
        public double GetNetTotal() { return this.netTotal; }
        public double GetVat() { return this.vat; }
        public double GetGrossTotal() { return this.grossTotal; }
        public int GetStatus() { return this.status; }
        public bool HasAttachment() { return this.attachment; }
    }
}
