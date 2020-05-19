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

namespace GetXml.Controllers
{
    public class HomeController : Controller
    {
        private ILoggerFactory _loggerFactory;
        private readonly DeviceRepository deviceRepository;
        public static string displayName = "(GMT+03:00) Russia Time Zone 2";
        public static string standardName = "Russia Time Zone 2";
        public static TimeSpan offset = new TimeSpan(03, 00, 00);
        public TimeZoneInfo moscow = TimeZoneInfo.CreateCustomTimeZone(standardName, offset, displayName, standardName);
        public HomeController(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _loggerFactory = loggerFactory;
            deviceRepository = new DeviceRepository(configuration);
            loggerFactory.AddFile(Path.Combine(Directory.GetCurrentDirectory(), "logger.txt"));
            deviceRepository = new DeviceRepository(configuration);
        }

        public IActionResult Index()
        {
            GetXmlData();            
            var terminals = deviceRepository.GetDevices();
            terminals = ConverDateToMoscowTime(terminals);
            return View(terminals);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        public async void GetXmlData()
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {
                var client = new HttpClient();
                string GETXML_PATH = "https://api.ar.digital/v4/devices/xml/M9as6DMRRGJqSVxcE9X58TM2nLbmR99w";
                //var responce = await client.GetAsync(GETXML_PATH);
                //var pageContent =  await responce.Content.ReadAsStringAsync();
                //var streamTaskA = client.GetStreamAsync(GETXML_PATH);
                var streamTaskA1 = await client.GetStringAsync(GETXML_PATH);
                XmlRootAttribute xRoot = new XmlRootAttribute();
                xRoot.ElementName = "xml";
                xRoot.DataType = "string";

                using (var reader = new StringReader(streamTaskA1))
                {
                    Xml devices = (Xml)(new XmlSerializer(typeof(Xml), xRoot)).Deserialize(reader);

                    foreach (Device d in devices.Devices)
                    {
                        if (deviceRepository.Get(d.Id) == null && d.Last_Online.Year == DateTime.Now.Year)
                        {
                            deviceRepository.Add(d);
                        }
                        if (d.Status == "offline" && (d.Last_Online > DateTime.MinValue) && d.Last_Online.Year == DateTime.Now.Year || d.Status == "playback" && (d.Last_Online > DateTime.MinValue) && d.Last_Online.Year == DateTime.Now.Year)
                        {
                            d.Hours_Offline = (DateTime.UtcNow - d.Last_Online).TotalHours;
                            var ts_new = TimeSpan.FromHours(d.Hours_Offline);
                            var h_new = System.Math.Floor(ts_new.TotalHours);
                            if (h_new < 0)
                            {
                                h_new = 0;
                            }
                            d.Hours_Offline = h_new;
                            if (h_new > 0)
                            {
                                var deviceFromDb = deviceRepository.Get(d.Id);
                                deviceFromDb.Last_Online = d.Last_Online;
                                deviceFromDb.Hours_Offline = d.Hours_Offline;
                                deviceFromDb.Status = d.Status;
                                deviceFromDb.Campaign_Name = d.Campaign_Name;
                                deviceRepository.Update(deviceFromDb);
                            }
                            else
                            {
                                var deviceFromDb = deviceRepository.Get(d.Id);
                                deviceFromDb.SumHours += deviceFromDb.Hours_Offline;
                                deviceFromDb.Hours_Offline = d.Hours_Offline;
                                deviceFromDb.Last_Online = d.Last_Online;
                                deviceFromDb.Status = d.Status;
                                deviceFromDb.Campaign_Name = d.Campaign_Name;
                                deviceRepository.Update(deviceFromDb);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                loggerF.LogError($"The path threw an exception {DateTime.Now} -- {e}");
                loggerF.LogWarning($"The path threw a warning {DateTime.Now} --{e}");
            }
        }

        public void CreateExcelReport()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                //var headerRow = new List<string[]>()
                //{
                //new string[] { "ID", "First Name", "Last Name", "DOB" }
                //};

                var TerminalList = deviceRepository.GetDevices();

                // Determine the header range (e.g. A1:D1)
                string headerRange = "A2:" + Char.ConvertFromUtf32(9 + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Status";
                worksheet.Cells[1, 4].Value = "Compaign Name";
                worksheet.Cells[1, 5].Value = "IP Address";
                worksheet.Cells[1, 6].Value = "Last Online";
                worksheet.Cells[1, 7].Value = "Address";
                worksheet.Cells[1, 8].Value = "Hours Offline";
                worksheet.Cells[1, 9].Value = "Sum Offline";

                // Popular header row data
                worksheet.Cells["A2"].LoadFromCollection(TerminalList);
                worksheet.Cells.Style.WrapText = true;
                worksheet.Column(6).Style.Numberformat.Format = "dd-MM-yyyy HH:mm";
                FileInfo excelFile = new FileInfo(@"D:\inetpub\vhosts\smartsoft83.com\httpdocs\report.xlsx");
                excel.SaveAs(excelFile);
            }
        }
            

            public List<string> ReadAddressesFromExcel() // make when a new device added for updating addresses//
            {
            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"D:\inetpub\vhosts\smartsoft83.com\httpdocs\Files\terminal_address.xlsx");

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
                            if ((worksheet.Cells[i, j].Value != null && j == 1 && worksheet.Cells[i, j + 7].Value != null) || (worksheet.Cells[i, j].Value != null && j == 8))
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

        public IActionResult PostAddressToDb()
        {
            var excelData = ReadAddressesFromExcel();
            excelData.RemoveRange(0, 2);
            for (int i = 0; i < excelData.Count - 1; i++)
            {
                var device = new Device(excelData[i], excelData[i + 1]);
                deviceRepository.AddAddress(device);
                i++;
            }
            return View("Privacy");
        }

        public FileResult Export()
        {
            CreateExcelReport();
            byte[] fileBytes = System.IO.File.ReadAllBytes(@"D:\inetpub\vhosts\smartsoft83.com\httpdocs\report.xlsx");
            string fileName = "terminals_report.xlsx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        public List<Device> ConverDateToMoscowTime(List<Device> listDevises)
        {
            foreach (var d in listDevises)
            {
                TimeZoneInfo moscowZone = TimeZoneInfo.FindSystemTimeZoneById("Russian Standard Time");
                d.Last_Online = TimeZoneInfo.ConvertTimeFromUtc(d.Last_Online, moscowZone);
                //string date_from = d.Last_Online.ToString("yyyy/MM/dd HH:mm");
            }
            return listDevises;
        }
    }
}


