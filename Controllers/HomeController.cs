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
        public List<string> excelDataAddress = new List<string>();
        public List<Tuple<double, string>> excelDataNotes = new List<Tuple<double, string>>();

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

            Task task3 = new Task(() => getHoursOffline());
            task3.Start();
            task3.Wait();

            var terminals = deviceRepository.GetDevices();
            terminals = ConverDateToMoscowTime(terminals);
            return View(terminals);
        }

        public IActionResult IndexSort(string sortOrder)
        {
            ViewData["NameSortParm"] = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            var compain = deviceRepository.GetDevices();
            switch (sortOrder)
            {
                case "name_desc":
                    compain = compain.OrderByDescending(s => s.Campaign_Name).ToList();
                    break;                
                
            }
            return View("Index",compain);
        }

        public async Task<IActionResult> Edit(double id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var device = deviceRepository.Get(id);

            if (device == null)
            {
                return NotFound();
            }

            return View(device);
        }

        [HttpPost, ActionName("Edit")]
        public async Task<IActionResult> EditPost(Device device)
        {
            if (device == null)
            {
                return NotFound();
            }

            deviceRepository.UpdateDevice(device);

            return View(device);
        }

        [DisableRequestSizeLimit]
        [HttpPost("Home")]
        public async Task<ViewResult> Index(IFormFile file)
        {
            long size = file.Length;
            
                if (size > 0)
                {
                // full path to file in temp location
                var filePath = Path.Combine(@"d:\Domains\smartsoft83.com\wwwroot\terminal\Files\", file.FileName); //we are using Temp file name just for the example. Add your own file path.
                //var filePath = Path.Combine(@"C:\Users\Timur\source\repos\GetXml\Files\", file.FileName); //we are using Temp file name just for the example. Add your own file path.

                using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }
                }

            PostAddressToDb();            
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

        public double getHoursOffline(Device device)
        {
            double h_new = 0;
            try
            {                
                var Hours_Offline = (DateTime.UtcNow - device.Last_Online).TotalHours;
                var ts_new = TimeSpan.FromHours(Hours_Offline);
                h_new = System.Math.Floor(ts_new.TotalHours);
                device.Hours_Offline = h_new;
                   
                if (h_new < 0)
                    h_new = 0;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }

            return h_new;
        }

        public void getHoursOffline()
        {
            double h_new = 0;
            try
            {
                foreach (var d in Devices.Devices.ToList())
                {
                    AddNewDevice(d);
                    //var Hours_Offline = (DateTime.UtcNow - d.Last_Online).TotalHours;
                    //var ts_new = TimeSpan.FromHours(Hours_Offline);
                    //h_new = System.Math.Floor(ts_new.TotalHours);                    
                    var deviceFromDb = deviceRepository.Get(d.Id);
                    if (d.Hours_Offline == deviceFromDb.Hours_Offline)
                        ChangeTerminalData(d);
                    else
                        OfflineHoursCount(d);
                }               
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }            
        }

        public async void GetXmlData()
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            loggerF.LogInformation($"Executed GetXmlData |======================================|  {DateTime.Now}");
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
                //XmlSerializer formatter = new XmlSerializer(typeof(Device));
                //using (FileStream fs = new FileStream(@"C:\Users\Timur\source\repos\GetXml\TestTerminalsData.xml", FileMode.OpenOrCreate))
                //{
                //    Xml devices = (Xml)(new XmlSerializer(typeof(Xml), xRoot)).Deserialize(fs);
                //    Devices = devices;
                //}
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
                    if (getHoursOffline(d) > 730 || (d.Status == "not attached") || (d.Last_Online < DateTime.MinValue) || (d.Last_Online.Year < 2020))
                    {
                        Devices.Devices.Remove(d);
                        deviceRepository.Delete(d.Id);
                    }
                }
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }            
        }        

        public void OfflineHoursCount(Device device)
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {       
                var deviceFromDb = deviceRepository.Get(device.Id);                    

                if (device.Hours_Offline == 48)
                {
                    deviceFromDb.SumHours += 48;
                    deviceFromDb.Hours_Offline = device.Hours_Offline;
                    deviceRepository.UpdateSumHourse(deviceFromDb);
                    deviceRepository.UpdateHoursOffline(deviceFromDb);
                    ChangeTerminalData(deviceFromDb);
                    loggerF.LogInformation($"Executed OfflineHoursCount Hours_Offline == 48 Device id = {deviceFromDb.Id}, Device Hours offline =  {deviceFromDb.Hours_Offline}|======================================|  {DateTime.Now}");
                }
                else if (device.Hours_Offline > 48 && deviceFromDb.SumHours >= 48)
                {
                    deviceFromDb.SumHours += 1;
                    deviceFromDb.Hours_Offline = device.Hours_Offline;
                    deviceRepository.UpdateSumHourse(deviceFromDb);
                    deviceRepository.UpdateHoursOffline(deviceFromDb);
                    ChangeTerminalData(deviceFromDb);
                    loggerF.LogInformation($"Executed OfflineHoursCount Hours_Offline > 48 Device id = {deviceFromDb.Id}, Device Hours offline =  {deviceFromDb.Hours_Offline}|=====================================|  {DateTime.Now}");
                }
                //else if (device.Hours_Offline > 48 && deviceFromDb.SumHours == 0)
                //{
                //    deviceFromDb.SumHours = Math.Floor(device.Hours_Offline / 24);
                //    deviceFromDb.Hours_Offline = device.Hours_Offline;
                //    deviceRepository.UpdateSumHourse(deviceFromDb);
                //    deviceRepository.UpdateHoursOffline(deviceFromDb);
                //    ChangeTerminalData(deviceFromDb.Id);
                //}
                else
                {                    
                    deviceRepository.UpdateHoursOffline(device);
                    ChangeTerminalData(deviceFromDb);
                }                     
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }       

        public void ChangeTerminalData(Device device)
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {
                deviceRepository.Update(device);                
            }
            catch (Exception ex)
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
                FileInfo excelFile = new FileInfo(@"d:\Domains\smartsoft83.com\wwwroot\terminal\report.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        public void ReadAddressesFromExcel() // make when a new device added for updating addresses//
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string DeviceNote = "";
            double DeviceId = 0;
            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"d:\Domains\smartsoft83.com\wwwroot\terminal\Files\Отчет по устройствам.xlsx");
            //byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\source\repos\GetXml\Files\Отчет по устройствам.xlsx");


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
                                excelDataAddress.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                            //if ((worksheet.Cells[i, j].Value != null && j == 1 && worksheet.Cells[i, j + 9].Value != null))
                            //{
                            //    DeviceId = Double.Parse(worksheet.Cells[i, j].Value.ToString());                                
                            //}
                            //if (worksheet.Cells[i, j].Value != null && j == 10)
                            //{
                            //    DeviceNote = worksheet.Cells[i, j].Value.ToString();

                            //}
                            //excelDataNotes.Add(new Tuple<double, string>(DeviceId,DeviceNote));
                        }
                    }
                }
            }            
        }

        public IActionResult PostAddressToDb()
        {
            ReadAddressesFromExcel();
            excelDataAddress.RemoveRange(0, 2);
            //excelDataNotes.RemoveRange(0, 2);

            for (int i = 0; i < excelDataAddress.Count - 1; i++)
            {
                var device = new Device(excelDataAddress[i], excelDataAddress[i + 1]);
                var dbDevice = deviceRepository.GetAddress(device.Name);
                if (dbDevice != null && dbDevice.Address != device.Address)
                    deviceRepository.UpdateAddress(device);
                else if (dbDevice != null && dbDevice.Address == "")
                    deviceRepository.UpdateAddress(device);
                i++;
            }

            //for (int i = 0; i < excelDataNotes.Count - 1; i++)
            //{

            //}
            return View("Privacy");
        }

        public FileResult Export()
        {
            CreateExcelReport();
            byte[] fileBytes = System.IO.File.ReadAllBytes(@"d:\Domains\smartsoft83.com\wwwroot\terminal\report.xlsx");
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

        public List<double> ImportTeamViewerCodesFromExcel() 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //create a list to hold all the values
            List<double> excelData = new List<double>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"d:\Domains\smartsoft83.com\wwwroot\Files\ICL (доступы TeamViewer).xlsx");

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
                        for (int j = worksheet.Dimension.Start.Column; j < 2; j++)
                        {                           
                            //    excelData.Add(worksheet.Cells[i, j].Value);
                            //worksheet.Cells.DataValidation.AddCustomDataValidation();
                            
                        }
                    }
                }
            }
            return excelData;
        }

        //public IActionResult PostTeamViewerCodesToDb()
        //{
        //    var excelData = ImportTeamViewerCodesFromExcel();
        //    excelData.RemoveRange(0, 2);
        //    for (int i = 0; i < excelData.Count - 1; i++)
        //    {
        //        var terminalsList = deviceRepository.GetDevices();
        //        foreach (var t in terminalsList)
        //        {
        //            if (t.Id == excelData[i])
        //        }




        //        var device = new Device(excelData[i], excelData[i + 1]);
        //        if (deviceRepository.GetAddress(device.Name) == null)
        //        {
        //            deviceRepository.AddAddress(device);
        //        }
        //        else if (deviceRepository.GetAddress(device.Name).Address != device.Address)
        //        {
        //            deviceRepository.UpdateAddress(device);
        //        }

        //        i++;
        //    }
        //    return View("Privacy");
        //}
    }
}


