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
            var markets = _repository.GetMarkets();
            foreach (var m in markets)
                _repository.InsertMarkets(m);
            return View();
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
        public void SortMarkets()
        {
            var noStockMarkets = _repository.GetMarkets();
            var savedMarkets = _repository.GetSavedMarkets();
            var newMarkets = noStockMarkets.Concat(savedMarkets).GroupBy(n => n.StoreId).
                Where(n => n.Count() == 1).Select(n => n.FirstOrDefault()).ToList();
            if (newMarkets.Count > 0)
            {
                foreach (var m in newMarkets)
                    _repository.InsertMarkets(m);
            }
            for (int i = 0; i < noStockMarkets.Count(); i ++)
                for (int j = 0; j < savedMarkets.Count(); j++)
                {
                    if (noStockMarkets[i].StoreId == savedMarkets[j].StoreId && DateTime.Now.Day != savedMarkets[j].TimeStamp.Day &&
                        DateTime.Today.Month != savedMarkets[j].TimeStamp.Month)
                        _repository.InsertMarkets(noStockMarkets[i]);
                }



        }
    }
}
