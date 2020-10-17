using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using FarmacyControl.Models;
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
        public List<Mrc> MrcList = new List<Mrc>();

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

            var data = ReadMrcFromExcel();



            return View("Index",data);
        }

        public List<string> ReadMrcFromExcel() // make when a new device added for updating addresses//
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DirectoryInfo di = new DirectoryInfo(@"C:\Users\t.lumelsky\source\repos\FarmacyControl\Files\");
            FileInfo[] files = di.GetFiles("*.xlsx");
            for (int i = 0; i < files.Length; i++)
            {
                long length = new System.IO.FileInfo(files[i].FullName).Length;
                if (length > 4000000)
                    files.SetValue(files[i], 0);
            }

            //read the Excel file as byte array
            //byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\source\repos\JnvlsList\Files\lp2020-06-09-1.xlsx");
            byte[] bin = System.IO.File.ReadAllBytes(files.First().FullName);
            //read the Excel file as byte array
            //byte[] bin = System.IO.File.ReadAllBytes(@"d:\Domains\smartsoft83.com\wwwroot\Files\Отчет по устройствам.xlsx");

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
            
            return excelData;
        }

        public ActionResult UploadToDb()
        {
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\t.lumelsky\source\repos\FarmacyControl\Files\");
            FileInfo[] files = di.GetFiles("*.xlsx");
            for (int i = 0; i < files.Length; i++)
            {
                long length = new System.IO.FileInfo(files[i].FullName).Length;
                if (length > 4000000)
                    files.SetValue(files[i], 0);
            }            
            byte[] bin = System.IO.File.ReadAllBytes(files.First().FullName);
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
            for (int i = 0; i < excelData.Count - 1; i++)
            {
                var MrcObj = new Mrc(excelData[i], excelData[i + 1]);
                MrcList.Add(MrcObj);
                Repository.WriteToDb(MrcObj);
                i++;
            }
            
            return View("Index");            
        }
    }
}
