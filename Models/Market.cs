using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.Models
{
    public class Market
    {       

        DateTime timeStamp;
        DateTimeOffset stockDate;

        public Market() { }
        public Market(Guid storeId, string storeName, string netName, string softwareName, DateTimeOffset stockDate, bool activeFl, bool reserveFl, bool stocksFl)
        {
            StoreId = storeId;
            StoreName = storeName;
            NetName = netName;
            SoftwareName = softwareName;
            if (stockDate != null)
                StockDate = stockDate;
            else
                StockDate = DateTimeOffset.Now;                

            ActiveFl = activeFl;
            ReserveFl = reserveFl;
            StocksFl = stocksFl;           
        }

        public Guid StoreId { get; set; }
        public string StoreName { get; set; }
        public string NetName { get; set; }
        public string SoftwareName { get; set; }
        public DateTimeOffset  StockDate
        {
            get
            {
                return stockDate;
            }
            set
            {
                if (value != null)
                    stockDate = value;
                else
                    stockDate = DateTimeOffset.MinValue;
                        
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
        public string Status { get; set; }
    }
}
