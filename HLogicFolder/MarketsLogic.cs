using FarmacyControl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.HLogicFolder
{
    public class MarketsLogic : IMarketsLogic
    {
        private IRepository _repository;

        public MarketsLogic(IRepository repository)
        {
            _repository = repository;
        }

        public void InsertNewMarkets()
        {
            var noStockMarkets = _repository.GetNewMarkets();
            var savedMarkets = _repository.GetSavedMarkets();
            if (savedMarkets.Count == 0)
            {
                var Markets = _repository.GetNewMarkets();
                foreach (var m in Markets)
                {
                    if (m.ReserveFl == false && m.ActiveFl == false && m.StocksFl == true)
                        m.Status = "in work";
                    m.Reason = "tech prbl";
                    _repository.InsertMarkets(m);
                }
            }

            var newMarkets = noStockMarkets.Concat(savedMarkets).GroupBy(n => n.StoreId).
            Where(n => n.Count() == 1).Select(n => n.FirstOrDefault()).ToList();
            if (newMarkets.Count > 0)
            {
                foreach (var m in newMarkets)
                {
                    if (m.ReserveFl == false && m.ActiveFl == false && m.StocksFl == true)
                        m.Status = "in work";
                    m.Reason = "tech prbl";
                    _repository.InsertMarkets(m);
                }
            }
        }
        public void UpdateCurrentListOfMarkets()
        {
            var noStockMarkets = _repository.GetNewMarkets();
            var savedMarkets = _repository.GetSavedMarkets();
            if (savedMarkets.Count == 0)
                InsertNewMarkets();

            for (var i = 0; i < savedMarkets.Count; i++)
            {
                for (var j = 0; j < noStockMarkets.Count; j++)
                {
                    // ищем магазины которые есть в обоих списках но ts новее и добавляем его с новой датой
                    if (savedMarkets[i].StoreId == noStockMarkets[j].StoreId && savedMarkets[i].TimeStamp.Day != noStockMarkets[j].TimeStamp.Day)
                    {
                        savedMarkets[i].TimeStamp = DateTime.Now;
                        _repository.InsertMarkets(savedMarkets[i]);
                        break;
                    }
                    // ищем магазины которые есть в обоих списках с одинаковой датой и пропускаем его
                    else if (savedMarkets[i].StoreId == noStockMarkets[j].StoreId && savedMarkets[i].TimeStamp.Day == noStockMarkets[j].TimeStamp.Day)
                    {
                        break;
                    }
                    // ищем новые магазины в старом листе и ксли их нет добавляем
                    var newMarket = savedMarkets.Select(m => m.StoreId).Contains(noStockMarkets[j].StoreId);
                    if (!newMarket)
                        _repository.InsertMarkets(savedMarkets[i]);

                }
                // ищем старые магазины в новом листе и если его уже нет и прошло <= 3 часа меняем статус
                var markeExist = noStockMarkets.Select(a => a.StoreId).Contains(savedMarkets[i].StoreId);
                if (!markeExist && savedMarkets[i].TimeStamp.Hour <= 3)
                {
                    savedMarkets[i].Status = "on-line";
                    savedMarkets[i].Reason = "> 24h";
                }
            }
        }
    }
}
