using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.Models
{
    public class Market
    {
        DateTime timeStamp;
        DateTime stockDate;
        public Guid StoreId { get; set; }
        public string StoreName { get; set; }
        public string NetName { get; set; }
        public string SoftwareName { get; set; }
        public DateTime StockDate {
            get {
                return stockDate;
            } 
            set {
                if (value <= DateTime.MinValue)
                    stockDate = value;
            } 
        }
        public bool ActiveFl { get; set; }
        public bool ReserveFl { get; set; }
        public bool StocksFl { get; set; }
        public DateTime TimeStamp {

            get {
                return timeStamp = DateTime.Now; 
            }          
        
        }
        public string Reason { get; set; }
    }
}
