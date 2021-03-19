using FarmacyControl;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.Controllers
{
    public class MarketController : Controller
    {
        private IRepository _repository;

        public MarketController(IRepository repository)
        {
            _repository = repository;
        }

        // GET: MarketController
        public ActionResult Index()
        {
            InsertNewMarkets();
            UpdateCurrentListOfMarkets();
            //var markets = _repository.GetNewMarkets();
            //foreach (var m in markets)
            //    _repository.InsertMarkets(m);
            var markets = _repository.GetSavedMarkets();
            return View(markets);
        }

        // GET: MarketController/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: MarketController/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MarketController/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MarketController/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: MarketController/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: MarketController/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: MarketController/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
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
        public void GetReport()
        {
            //var markets = _repository.GetNewMarkets();

        }
    }
}
