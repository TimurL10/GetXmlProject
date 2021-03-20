using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using FarmacyControl.Models;
using GetXml.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace FarmacyControl.Controllers
{
    public class MrcController : Controller
    {

        private Repository Repository;
        public List<string> excelData = new List<string>();
        public static List<Mrc> MrcList = new List<Mrc>();

        public MrcController(IConfiguration _configuration)
        {
            Repository = new Repository(_configuration);

        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("Index")]
        public async Task<ViewResult> IndexPost(IFormFile file)
        {
            List<string> data;
            try
            {
                long size = file.Length;

                if (size > 0)
                {
                    // full path to file in temp location
                    var filePath = Path.Combine(@"C:\Users\t.lumelsky\source\repos\FarmacyControl\Files\", file.FileName); //we are using Temp file name just for the example. Add your own file path.

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }
                }

                data = ReadMrcFromExcel();
            }
            catch
            {
                ErrorViewModel errorViewModel = new ErrorViewModel();
                errorViewModel.RequestId = "No File Was Choosen";
                return View("Error", errorViewModel);
            }

            return View("Index", data);
        }

        public List<string> ReadMrcFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DirectoryInfo di = new DirectoryInfo(@"C:\Users\t.lumelsky\source\repos\FarmacyControl\Files\");
            FileInfo[] files = di.GetFiles("*.xlsx");

            byte[] bin = System.IO.File.ReadAllBytes(files.First().FullName);

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells != null)
                            {
                                //if (worksheet.Cells["A"])
                                excelData.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                        }
                    }
                }
            }
            // Creating Mrc Objects
            for (int i = 0; i < excelData.Count - 1; i++)
            {
                var MrcObj = new Mrc(excelData[i], excelData[i + 1]);
                MrcList.Add(MrcObj);
                i++;
            }

            System.IO.File.Delete(files.First().FullName);

            return excelData;
        }

        public ActionResult UploadToDb()
        {
            List<Mrc> NewMrcListForInsert = new List<Mrc>();
            Mrc sameProduct = new Mrc();
            var listCurrentmrc = Repository.GetMrc();

            foreach (var a in listCurrentmrc) // update
            {
                sameProduct = MrcList.Find(x => x.Nnt == a.Nnt);
                if (sameProduct != null)
                {
                    a.Price = sameProduct.Price;
                    //Repository.UpdateDb(a);
                }

            }
            foreach (var a in MrcList) // insert
            {
                sameProduct = listCurrentmrc.Find(x => x.Nnt == a.Nnt);
                if (sameProduct == null)
                {
                    NewMrcListForInsert.Add(a);
                    //Repository.InsertDb(a);
                }
            }

            MrcList.Clear();
            return View("Index");
        }
    }
}
