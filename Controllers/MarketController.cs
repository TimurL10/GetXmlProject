using FarmacyControl;
using GetXml.HLogicFolder;
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
        private IMarketsLogic _marketsLogic;

        public MarketController(IMarketsLogic marketsLogic)
        {
            _marketsLogic = marketsLogic;
        }

        // GET: MarketController
        public ActionResult Index()
        {                  
            
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

        public void InsertNewMarkets()
        {
            _marketsLogic.InsertNewMarkets(); 
        }

        public void UpdateCurrentListOfMarkets()
        {
            _marketsLogic.UpdateCurrentListOfMarkets();            
        }

        public void GetReport()
        {
            //var markets = _repository.GetNewMarkets();

        }
    }
}
