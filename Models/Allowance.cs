using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace JnvlsList.Models
{
    public class Allowance
    {
        public string Oblast { get; set; }
        public double Till50Whole { get; set; }
        public double Till500Whole { get; set; }
        public double After500Whole { get; set; }
        public double Till50Retail { get; set; }
        public double Till500Retail { get; set; }
        public double After500Retail { get; set; }

    }
}
