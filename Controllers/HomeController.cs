using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using GetXml.Models;
using System.Net.Http;
using System.Xml.Serialization;
using System.IO;
using System.Xml;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;

namespace GetXml.Controllers
{
    public class HomeController : Controller
    {    
        public IActionResult Index()
        {            
            return View();
        }   

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }       
              
    }
}


