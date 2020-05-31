using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using JnvlsList.Model;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;
using JnvlsList.Models;
using Microsoft.AspNetCore.Http;

namespace JnvlsList.Controllers
{
    public class DrugController : Controller
    {
        List<String> Regions = new List<string>(new string[] { "Владимирская область","Республика Башкортостан","Ивановская область","Иркутская область","Калининградская область","Краснодарский край","Красноярский край",
        "г.Москва","Московская область","Нижегородская область","Новгородская область","Омская область","Орловская область","Приморский край","Ростовская область","Республика Саха (Якутия)","Свердловская область","Тверская область","Республика Хакасия"});
        List<Allowance> ListAllowance = new List<Allowance>();
        public static string header;
        public List<Drug> VladimirskayaObl = new List<Drug>();
        List<Drug> Bashkiriya = new List<Drug>();
        List<Drug> IvanovskayaObl = new List<Drug>();
        List<Drug> IrkutskayaObl = new List<Drug>();
        List<Drug> KaliningradskayaObl = new List<Drug>();
        List<Drug> Krasnodarskii = new List<Drug>();
        List<Drug> Krasnoyarskii = new List<Drug>();
        List<Drug> Moskva = new List<Drug>();
        List<Drug> MO = new List<Drug>();
        List<Drug> NijnegorodskayaObl = new List<Drug>();
        List<Drug> NovgorodskayaObl = new List<Drug>();
        List<Drug> OmskayaObl = new List<Drug>();
        List<Drug> OrlovskayaObl = new List<Drug>();
        List<Drug> Primorskii = new List<Drug>();
        List<Drug> RostovskayaObl = new List<Drug>();
        List<Drug> Saha = new List<Drug>();
        List<Drug> Sverdlovskaya = new List<Drug>();
        List<Drug> tverskaya = new List<Drug>();
        List<Drug> Hakasiya = new List<Drug>();

        public IActionResult Index()
        {
            //ReadDrugsFromExcel();            
            //ReadDotationsFromExcel();
            //CalculateDotation();
            return View();
        }

        [HttpPost("Drug")]
        public async Task<IActionResult> Index(List<IFormFile> files)
        {
            long size = files.Sum(f => f.Length);

            var filePaths = new List<string>();
            foreach (var formFile in files)
            {
                if (formFile.Length > 0)
                {
                    // full path to file in temp location
                    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Files");
                    //GetFullPath(@"C:\Users\Timur\source\repos\JnvlsList\Files"); //we are using Temp file name just for the example. Add your own file path.
                    filePaths.Add(filePath);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await formFile.CopyToAsync(stream);
                    }
                }
            }

            // process uploaded files
            // Don't rely on or trust the FileName property without validation.

            return Ok(new { count = files.Count, size, filePaths });
        }




        //[HttpPost]
        //    public async Task<IActionResult> UploadFile(IFormFile file)
        //    {
        //        if (file != null)
        //            return Content("file not selected");

        //        string path = @"C:\Users\Timur\source\repos\JnvlsList\Files\" + file.FileName;

        //        using (var stream = new FileStream(path, FileMode.Create))
        //        {
        //            await file.CopyToAsync(stream);
        //        }

        //        return RedirectToAction("Index");
        //    }



        public void ReadDrugsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\source\repos\JnvlsList\Files\lp2020-05-22-1.xlsx");

            //or if you use asp.net, get the relative path
            //byte[] bin = File.ReadAllBytes(Server.MapPath("ExcelDemo.xlsx"));

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))

            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet  in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= 11; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value == null )
                            {
                                excelData.Add(null);                                
                            }
                            else
                            {                             
                                excelData.Add(char.ToUpper(worksheet.Cells[i, j].Value.ToString()[0]) + worksheet.Cells[i, j].Value.ToString().Substring(1));                              
                            }
                        }
                    }
                }
            }            
            DeserializeExcel(excelData);
        }

        public void DeserializeExcel(List<string> list)
        {
            List<Drug> DrugList = new List<Drug>();
            header = list.First().ToString();

            if (list != null)
            {
                for (int i = 33; i < list.Count-10; i += 11)
                {
                    var drug = new Drug(list[i + 1], list[i], list[i + 2], list[i + 3], list[i + 4], list[i + 5], list[i + 7], list[i + 8], list[i + 9], list[i + 10],list[i + 6]);
                    DrugList.Add(drug);
                }
            }
            List<Drug> SortedList = DrugList.OrderBy(o => o.Name).ToList();
            VladimirskayaObl = new List<Drug>(SortedList);
             Bashkiriya = new List<Drug>(SortedList);
            IvanovskayaObl = new List<Drug>(SortedList);
             IrkutskayaObl = new List<Drug>(SortedList);
            KaliningradskayaObl = new List<Drug>(SortedList);
            Krasnodarskii = new List<Drug>(SortedList);
            Krasnoyarskii = new List<Drug>(SortedList);
            Moskva = new List<Drug>(SortedList);
            MO = new List<Drug>(SortedList);
            NijnegorodskayaObl = new List<Drug>(SortedList);
            NovgorodskayaObl = new List<Drug>(SortedList);
            OmskayaObl = new List<Drug>(SortedList);
            OrlovskayaObl = new List<Drug>(SortedList);
            Primorskii = new List<Drug>(SortedList);
            RostovskayaObl = new List<Drug>(SortedList);
            Saha = new List<Drug>(SortedList);
            Sverdlovskaya = new List<Drug>(SortedList);
            tverskaya = new List<Drug>(SortedList);
            Hakasiya = new List<Drug>(SortedList);
        }

        public void ReadDotationsFromExcel()
        {
            List<Drug> listDrugs = new List<Drug>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\source\repos\JnvlsList\Files\Allowances.xlsx");

            //or if you use asp.net, get the relative path
            //byte[] bin = File.ReadAllBytes(Server.MapPath("ExcelDemo.xlsx"));

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
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                foreach (var r in Regions)
                                {
                                    if (worksheet.Cells[i, j].Value.ToString() == r)
                                    {
                                        Allowance allowance = new Allowance();
                                        allowance.Oblast = worksheet.Cells[i, j].Value.ToString();
                                        if (worksheet.Cells[i + 1, j].Value.ToString() == "1 зона")
                                        {
                                            if (worksheet.Cells[i + 2, j + 1].Value.ToString().Contains("*"))
                                            {
                                                worksheet.Cells[i + 2, j + 1].Value = worksheet.Cells[i + 2, j + 1].Value.ToString().Remove(0,3);
                                                worksheet.Cells[i + 2, j + 1].Value = worksheet.Cells[i + 2, j + 1].Value.ToString().Remove(worksheet.Cells[i + 2, j + 1].Value.ToString().Length - 1, 1);
                                                worksheet.Cells[i + 3, j + 1].Value = worksheet.Cells[i + 3, j + 1].Value.ToString().Remove(0, 3);
                                                worksheet.Cells[i + 3, j + 1].Value = worksheet.Cells[i + 3, j + 1].Value.ToString().Remove(worksheet.Cells[i + 3, j + 1].Value.ToString().Length - 1, 1);
                                                worksheet.Cells[i + 4, j + 1].Value = worksheet.Cells[i + 4, j + 1].Value.ToString().Remove(0, 3);
                                                worksheet.Cells[i + 4, j + 1].Value = worksheet.Cells[i + 4, j + 1].Value.ToString().Remove(worksheet.Cells[i + 4, j + 1].Value.ToString().Length - 1, 1);


                                            }
                                            allowance.Till50Whole = (worksheet.Cells[i + 2, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 2, j + 1].Value.ToString()) : 0;
                                            allowance.Till50Retail = (worksheet.Cells[i + 2, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 2, j + 3].Value.ToString()) : 0;
                                            allowance.Till500Whole = (worksheet.Cells[i + 3, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 3, j + 1].Value.ToString()) : 0;
                                            allowance.Till500Retail = (worksheet.Cells[i + 3, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 3, j + 3].Value.ToString()) : 0;
                                            allowance.After500Whole = (worksheet.Cells[i + 4, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 4, j + 1].Value.ToString()) : 0;
                                            allowance.After500Retail = (worksheet.Cells[i + 4, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 4, j + 3].Value.ToString()) : 0;
                                        }
                                        else
                                        {
                                            allowance.Till50Whole = (worksheet.Cells[i + 1, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 1, j + 1].Value.ToString()) : 0;
                                            allowance.Till50Retail = (worksheet.Cells[i + 1, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 1, j + 3].Value.ToString()) : 0;
                                            allowance.Till500Whole = (worksheet.Cells[i + 2, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 2, j + 1].Value.ToString()) : 0;
                                            allowance.Till500Retail = (worksheet.Cells[i + 2, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 2, j + 3].Value.ToString()) : 0;
                                            allowance.After500Whole = (worksheet.Cells[i + 3, j + 1].Value != null) ? Double.Parse(worksheet.Cells[i + 3, j + 1].Value.ToString()) :0;
                                            allowance.After500Retail = (worksheet.Cells[i + 3, j + 3].Value != null) ? Double.Parse(worksheet.Cells[i + 3, j + 3].Value.ToString()) :0;
                                        }
                                        ListAllowance.Add(allowance);
                                    }
                                }
                              }
                           }                           
                        }
                    }
                }
        }

        public void CreateExcelReport(List<Drug> listDrugs, string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Drug> listByFirstW = new List<Drug>();
            using (ExcelPackage excel = new ExcelPackage())
            {               
                    for (int i = 0; i < listDrugs.Count-1; i++)
                    {                   
                            listByFirstW.Add(listDrugs[i]);

                        if (listDrugs[i].Name.First() != listDrugs[i + 1].Name.First())
                        {
                            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(listDrugs[i].Name.First().ToString());
                            worksheet.Cells[1, 1].Value = header;
                            worksheet.Cells["A2"].Value = "Торговое наименование лекарственного препарата";
                            worksheet.Cells["B2"].Value = "МНН";
                            worksheet.Cells[2, 3].Value = "Лекарственная форма, дозировка, упаковка (полная)";
                            worksheet.Cells[2, 4].Value = "Владелец РУ/производитель/упаковщик/Выпускающий контроль";
                            worksheet.Cells[2, 5].Value = "Код АТХ";
                            worksheet.Cells[2, 6].Value = "Количество в потреб.упаковке";
                            worksheet.Cells[2, 7].Value = "Цена указана для первич. упаковки";
                            worksheet.Cells[2, 8].Value = "№ РУ";
                            worksheet.Cells[2, 9].Value = "Дата регистрации цены(№ решения)";
                            worksheet.Cells[2, 10].Value = "Штрих-код (EAN13)";
                            worksheet.Cells[2, 11].Value = "Предельная цена руб. без НДС";
                            worksheet.Cells[2, 12].Value = "Предельная цена руб. с НДС";
                            worksheet.Cells[2, 13].Value = "Предельная оптовая надбавка, руб.*";
                            worksheet.Cells[2, 14].Value = "Предельная оптовая цена, руб., (без НДС)*";
                            worksheet.Cells[2, 15].Value = "Предельная оптовая цена руб., (с НДС)*";
                            worksheet.Cells[2, 16].Value = "Предельная розничная надбавка, руб.*";
                            worksheet.Cells[2, 17].Value = "Предельная розничная цена,  руб. (без НДС)*";
                            worksheet.Cells[2, 18].Value = "Предельная розничная цена,  руб. (с НДС)*";
                            worksheet.Cells.Style.WrapText = true;
                            worksheet.Cells["A3"].LoadFromCollection<Drug>(listByFirstW);
                            listByFirstW.Clear();
                        }
                    }                             
                FileInfo excelFile = new FileInfo($@"C:\Users\Timur\source\repos\JnvlsList\Files\{fileName}.xlsx");                
                excel.SaveAs(excelFile);
            }           
        }


        public void CalculateDotation()
        {
            foreach( var a in ListAllowance)
            {
                if (a.Oblast == "Владимирская область")
                {
                    foreach (var d in VladimirskayaObl)
                    {                        
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance,2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance),2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance,2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10),2);
                    }
                        CreateExcelReport(VladimirskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Республика Башкортостан")
                {
                    foreach (var d in Bashkiriya)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Bashkiriya, a.Oblast);
                }
                else if (a.Oblast == "Ивановская область")
                {
                    foreach (var d in IvanovskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(IvanovskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Иркутская область")
                {
                    foreach (var d in IrkutskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(IrkutskayaObl, a.Oblast);

                }
                else if (a.Oblast == "Калининградская область")
                {
                    foreach (var d in KaliningradskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(KaliningradskayaObl, a.Oblast);

                }
                else if (a.Oblast == "Краснодарский край")
                {
                    foreach (var d in Krasnodarskii)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Krasnodarskii, a.Oblast);
                }
                else if (a.Oblast == "Красноярский край")
                {
                    foreach (var d in Krasnoyarskii)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Krasnoyarskii, a.Oblast);
                }
                else if (a.Oblast == "г.Москва")
                {
                    foreach (var d in Moskva)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Moskva, a.Oblast);
                }
                else if (a.Oblast == "Московская область")
                {
                    foreach (var d in MO)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(MO, a.Oblast);
                }
                else if (a.Oblast == "Нижегородская область")
                {
                    foreach (var d in NijnegorodskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(NijnegorodskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Новгородская область")
                {
                    foreach (var d in NovgorodskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(NovgorodskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Омская область")
                {
                    foreach (var d in OmskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(OmskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Орловская область")
                {
                    foreach (var d in OrlovskayaObl)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(OrlovskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Приморский край")
                {
                    foreach (var d in Primorskii)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Primorskii, a.Oblast);
                }
                else if (a.Oblast == "Ростовская область")
                {
                    foreach (var d in RostovskayaObl)
                    {
                        {
                            d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                            if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                                d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                            else if (d.PriceNoNds <= 500)
                                d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                            else
                                d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                            d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                            d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                            d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                            if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                                d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                            else if (d.PriceNoNds < 500)
                                d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                            else
                                d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                            d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                            d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                            d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                        }
                    }
                    CreateExcelReport(RostovskayaObl, a.Oblast);
                }
                else if (a.Oblast == "Республика Саха (Якутия)")
                {
                    foreach (var d in Saha)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Saha, a.Oblast);
                }
                else if (a.Oblast == "Свердловская область")
                {
                    foreach (var d in Sverdlovskaya)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Sverdlovskaya, a.Oblast);
                }
                else if (a.Oblast == "Тверская область")
                {
                    foreach (var d in tverskaya)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(tverskaya, a.Oblast);
                }
                else if (a.Oblast == "Республика Хакасия")
                {
                    foreach (var d in Hakasiya)
                    {
                        d.PriceWithNds = Math.Round(((d.PriceNoNds / 100) * 10) + d.PriceNoNds, 2);
                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*17%;ЕСЛИ(K2<=500;K2*14%;K2*8,5%))
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till50Whole;
                        else if (d.PriceNoNds <= 500)
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.Till500Whole;
                        else
                            d.WholesaleAllowance = (d.PriceNoNds / 100) * a.After500Whole;
                        d.WholesaleAllowance = Math.Round(d.WholesaleAllowance, 2);

                        d.WholesalePriceNoNds = Math.Round(d.PriceNoNds + d.WholesaleAllowance, 2); //оптовая без ндс
                        d.WholesalePriceWithNds = Math.Round((((d.PriceNoNds + d.WholesaleAllowance) / 100) * 10) + (d.PriceNoNds + d.WholesaleAllowance), 2); // (K2+M2)+((K2+M2)*10%)

                        if (d.PriceNoNds < 51) //ЕСЛИ(K2<51;K2*31%;ЕСЛИ(K2<=500;K2*25%;K2*19%))
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till50Retail;
                        else if (d.PriceNoNds < 500)
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.Till500Retail;
                        else
                            d.RetailAllowance = (d.PriceNoNds / 100) * a.After500Retail;
                        d.RetailAllowance = Math.Round(d.RetailAllowance, 2);
                        d.RetailPriceNoNds = Math.Round(d.WholesalePriceNoNds + d.RetailAllowance, 2);
                        d.RetailPriceWithNds = Math.Round((d.WholesalePriceNoNds + d.RetailAllowance) + (((d.WholesalePriceNoNds + d.RetailAllowance) / 100) * 10), 2);
                    }
                    CreateExcelReport(Hakasiya, a.Oblast);
                }
            }
        }
  }









        //public void Convert(String filesFolder)
        //{
        //    var file = Directory.GetFiles(filesFolder);

        //    var app = new Microsoft.Office.Interop.Excel.Application();
        //    var wb = app.Workbooks.Open(file);
        //    wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
        //    wb.Close();
        //    app.Quit();
        //}


 }
