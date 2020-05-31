using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace JnvlsList.Model
{
    public class Drug
    {        
        public string Name { get; set; }
        public string Mnn { get; set; }
        public string DrugDetails { get; set; }
        public string DrugFactory { get; set; }
        public string AthCode { get; set; }
        public int NumInPacket { get; set; }
        public string RecomendedPrice { get; set; }
        public string RuNum { get; set; }
        public string PriceRegDate { get; set; }
        public string BarCode { get; set; }
        public double PriceNoNds { get; set; }
        public double PriceWithNds { get; set; } // PriceNoNds += 10%
        public double WholesaleAllowance { get; set; } // оптовая надбавка
        public double WholesalePriceNoNds { get; set; }
        public double WholesalePriceWithNds { get; set; }
        public double RetailAllowance { get; set; } // розничная надбавка
        public double RetailPriceNoNds { get; set; }
        public double RetailPriceWithNds { get; set; }

        public Drug(string name, string mnn, string drugDetails, string drugFactory, string athCode, string numInPacket, string recomendedPrice, string ruNum, string priceRegDate, string barCode, string priceNoNds)
        {           
            Name = name;
            Mnn = mnn;
            DrugDetails = drugDetails;
            DrugFactory = drugFactory;
            AthCode = athCode;
            if (numInPacket != null)
                NumInPacket = Int32.Parse(numInPacket);
            else
                NumInPacket = 0;            
            RecomendedPrice = recomendedPrice;
            RuNum = ruNum;
            PriceRegDate = priceRegDate;
            BarCode = barCode;
            if (priceNoNds != null)
                PriceNoNds = Double.Parse(priceNoNds);
            else
                PriceNoNds = 0;
        }
    }
}
