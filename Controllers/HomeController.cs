using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using GetXml.Models;
using System.Net.Http;
using System.Xml.Serialization;
using System.IO;
using System.Xml;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using Microsoft.AspNetCore.Diagnostics;

namespace GetXml.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private ILoggerFactory _loggerFactory;
        private readonly DeviceRepository deviceRepository;
        public HomeController(IConfiguration configuration, ILoggerFactory loggerFactory)
        {
            //_logger = logger;
            _loggerFactory = loggerFactory;
            deviceRepository = new DeviceRepository(configuration);
            loggerFactory.AddFile(Path.Combine(Directory.GetCurrentDirectory(), "logger.txt"));

            
        }
        
        public IActionResult Index()
        {
            GetXmlData();
            var terminals = deviceRepository.GetDevices();
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
                string GETXML_PATH = "https://api.ar.digital/v4/devices/xml/M9as6DMRRGJqSVxcE9X58TM2nLb";
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
                        if (deviceRepository.Get(d.Id) == null)
                        {
                            deviceRepository.Add(d);
                        }
                        if (d.Status == "offline" || d.Status == "playback")
                        {
                            d.Hours_Offline = (DateTime.UtcNow - d.Last_Online).TotalHours;
                            var ts = TimeSpan.FromHours(d.Hours_Offline);
                            var h = System.Math.Floor(ts.TotalHours);
                            d.Hours_Offline = h;
                            deviceRepository.Update(d);
                            Console.WriteLine($"Имя: {d.Name} Status: {d.Status} Id: {d.Id} hours_offline: {d.Hours_Offline} last_online: {d.Last_Online} Compaign {d.Campaign_Name}");
                        }
                    }
                }
            }            
            catch (Exception e)
            {
                loggerF.LogError($"The path threw an exception {e}");
                loggerF.LogWarning($"The path threw a warning {e}");                
            }
            CreateExcelReport();
           //PostAddressToDb(); 
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
                string headerRange = "A2:" + Char.ConvertFromUtf32(8 + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Status";
                worksheet.Cells[1, 4].Value = "Compaign Name"; 
                worksheet.Cells[1, 5].Value = "IP Address";
                worksheet.Cells[1, 6].Value = "Last Online"; 
                worksheet.Cells[1, 7].Value = "Hours Offline";
                worksheet.Cells[1, 8].Value = "Address";


                // Popular header row data
                worksheet.Cells["A2"].LoadFromCollection(TerminalList);               
                worksheet.Cells.Style.WrapText = true;
                worksheet.Column(6).Style.Numberformat.Format = "dd-MM-yyyy HH:mm";
                FileInfo excelFile = new FileInfo(@"C:\Users\Timur\Documents\test.xlsx");                
                excel.SaveAs(excelFile);
            }
        }

        public List<string> ReadAddressesFromExcel() // make when a new device added for updating addresses//
        {
            //create a list to hold all the values
            List<string> excelData = new List<string>();

            //read the Excel file as byte array
            byte[] bin = System.IO.File.ReadAllBytes(@"C:\Users\Timur\Documents\terminals.xlsx");
            
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
                            if ((worksheet.Cells[i, j].Value != null && j == 1 && worksheet.Cells[i, j+7].Value != null) || (worksheet.Cells[i, j].Value != null && j == 8))
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
            for (int i = 0; i < excelData.Count-1; i ++)
            {
                var device = new Device(excelData[i], excelData[i + 1]);                
                deviceRepository.AddAddress(device);
                i++;
            }
            return View("Privacy");
        }

        public FileResult Export()
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(@"C:\Users\Timur\Documents\test.xlsx");
            string fileName = "terminals_report.xlsx";
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
    } 
}


