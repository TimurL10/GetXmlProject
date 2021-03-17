using FarmacyControl.Models;
using GetXml.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FarmacyControl
{
    public interface IRepository
    {
        public List<Mrc> GetMrc();
        public void UpdateDb(Mrc mrcs);
        public void InsertDb(Mrc mrcs);
        public List<Market> GetNewMarkets();
        public void InsertMarkets(Market market);
        public List<Market> GetSavedMarkets();
        public void UpdateMarkets(Market market);
        public void DeleteActiveMarket(Market market);
    }
}
