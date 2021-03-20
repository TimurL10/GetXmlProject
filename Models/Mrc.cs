using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Threading.Tasks;

namespace FarmacyControl.Models
{
    public class Mrc
    {
        public int Nnt { get; set; }
        public int Price { get; set; }

        public Mrc(string nnt, string price)
        {
            this.Nnt = Int32.Parse(nnt);
            this.Price = (int)Double.Parse(price);
        }
        public Mrc()
        {

        }
    }
}
