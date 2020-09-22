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
        public Xml Devices;
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
            Task task1 = new Task(() => GetXmlData());
            task1.Start();
            task1.Wait();
            Task task2 = new Task(() => FilterDevices());
            task2.Start();
            task2.Wait();
            Task task3 = new Task(() => OfflineHoursCount());
            task3.Start();
            task3.Wait();
            var terminals = deviceRepository.GetDevices();
            terminals = ConverDateToMoscowTime(terminals);
            return View(terminals);
        }

        [DisableRequestSizeLimit]
        [HttpPost("Home")]
        public async Task<ViewResult> Index(IFormFile file)
        {
            long size = file.Length;
            
                if (size > 0)
                {
                    // full path to file in temp location
                    var filePath = Path.Combine(@"C:\Users\Timur\source\repos\GetXml\Reports", file.FileName); //we are using Temp file name just for the example. Add your own file path.

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }
                }
            
            // process uploaded files
            // Don't rely on or trust the FileName property without validation.
            return View("Privacy");
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

        public double getHoursOffline(double deviceId)
        {
            double h_new = 0;
            foreach (var d in Devices.Devices)
            {
                if (d.Id == deviceId)
                {
                    var Hours_Offline = (DateTime.UtcNow - d.Last_Online).TotalHours;
                    var ts_new = TimeSpan.FromHours(Hours_Offline);
                    h_new = System.Math.Floor(ts_new.TotalHours);
                }
            }
            if (h_new < 0)            
                h_new = 0;
            
            return h_new;
        }

        public async void GetXmlData()
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {
                var client = new HttpClient();
                string GETXML_PATH = "https://api.ar.digital/v4/devices/xml/M9as6DMRRGJqSVxcE9X58TM2nLbmR99w";
                var streamTaskA1 = client.GetStringAsync(GETXML_PATH).Result;
                XmlRootAttribute xRoot = new XmlRootAttribute();
                xRoot.ElementName = "xml";
                xRoot.DataType = "string";

                using (var reader = new StringReader(streamTaskA1))
                {
                    Xml devices = (Xml)(new XmlSerializer(typeof(Xml), xRoot)).Deserialize(reader);
                    Devices = devices;
                }
            }
            catch (Exception e)
            {
                loggerF.LogError($"The path threw an exception {DateTime.Now} -- {e}");
                loggerF.LogWarning($"The path threw a warning {DateTime.Now} --{e}");
            }
        }

        public void FilterDevices()
        {
            try
            {
                foreach (var d in Devices.Devices.ToList())
                {                   
                    if ( (getHoursOffline(d.Id) > 730) || (d.Status == "not attached") || (d.Last_Online < DateTime.MinValue))
                    {
                        Devices.Devices.Remove(d);
                    }
                }
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            
        }

        public void OfflineHoursCount()
        {
            try
            {
                foreach (var d in Devices.Devices.ToList())
                {
                    AddNewDevice(d);
                    var deviceFromDb = deviceRepository.Get(d.Id);
                    if (!deviceFromDb.Hours_Offline.Equals(getHoursOffline(d.Id)))
                        deviceFromDb.Hours_Offline = getHoursOffline(d.Id);

                    if (deviceFromDb.Hours_Offline == 48 && deviceFromDb.SumHours < 2)
                    {
                        deviceFromDb.SumHours += Math.Round(deviceFromDb.Hours_Offline / 24, 0);
                        deviceRepository.Update(deviceFromDb);
                        ChangeTerminalData(deviceFromDb.Id);
                    }

                    else if (deviceFromDb.Hours_Offline > 48 && deviceFromDb.SumHours >= 2 && (!deviceFromDb.Hours_Offline.Equals(getHoursOffline(d.Id))))
                    {
                        if ((deviceFromDb.Hours_Offline - deviceFromDb.SumHours * 24) == 24)
                        {
                            deviceFromDb.SumHours += 1;
                            deviceRepository.Update(deviceFromDb);
                            ChangeTerminalData(deviceFromDb.Id);
                        }
                    }
                    else
                        ChangeTerminalData(deviceFromDb.Id); // we can change to send device instead id
                }
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            
        }

        public void ChangeTerminalData(double deviceId)
        {
            try
            {
                var deviceFromDb = deviceRepository.Get(deviceId);
                foreach (var d in Devices.Devices.ToList())
                {
                    if (d.Id == deviceId)
                    {
                        deviceFromDb.Hours_Offline = getHoursOffline(d.Id);
                        deviceFromDb.Last_Online = d.Last_Online;
                        deviceFromDb.Status = d.Status;
                        deviceFromDb.Campaign_Name = d.Campaign_Name;
                        deviceFromDb.Address = d.Address;
                        deviceRepository.Update(deviceFromDb); 
                    }
                }

            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }          
        }

        public bool AddNewDevice(Device device)
        {
            bool flag = false;
            try
            {
                if (deviceRepository.Get(device.Id) == null)
                {
                    deviceRepository.Add(device);
                    return flag = true;
                }
                else
                    return flag = false;

            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return flag;            
        }

        public void CreateExcelReport()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

               
                var TerminalList = deviceRepository.GetDevices();

                // Determine the header range (e.g. A1:D1)
                string headerRange = "A2:" + Char.ConvertFromUtf32(8 + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "";
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
                FileInfo excelFile = new FileInfo(@"C:\Users\Timur\Documents\report.xlsx");
                excel.SaveAs(excelFile);
            }
        }


        public List<string> ReadAddressesFromExcel() // make when a new device added for updating addresses//
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\source\repos\GetXml\Reports\Отчет по устройствам.xlsx");

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
                            if ((worksheet.Cells[i, j].Value != null && j == 2 && worksheet.Cells[i, j + 7].Value != null) || (worksheet.Cells[i, j].Value != null && j == 9))
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
                if (deviceRepository.GetAddress(device.Name).Address != device.Address)
                {
                    deviceRepository.UpdateAddress(device);
                }
                else if (deviceRepository.GetAddress(device.Name) == null)
                {
                    deviceRepository.AddAddress(device);
                }
                i++;
            }
            return View("Privacy");
        }

        public FileResult Export()
        {
            CreateExcelReport();
            byte[] fileBytes = System.IO.File.ReadAllBytes(@"C:\Users\Timur\Documents\report.xlsx");
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


