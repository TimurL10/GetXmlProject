using GetXml.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace GetXml
{
    public class HLogic : IHLogic
    {
        private ILoggerFactory _loggerFactory;
        public Xml Devices;
        public IDeviceRepository _deviceRepository;        

        public HLogic(ILoggerFactory loggerFactory, IConfiguration configuration, IDeviceRepository deviceRepository)
        {
            _loggerFactory = loggerFactory;
            _deviceRepository = deviceRepository;
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

        public void SaveActivity()
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {
                foreach (var d in Devices.Devices)
                {

                     if (d.Status == "offline")
                        _deviceRepository.PostActivity(d.Id, DateTime.UtcNow, 0);
                }
            }
            catch(Exception ex)
            {
                loggerF.LogError($"The path threw an exception SaveActivity {DateTime.Now} =========== {ex.Message}");
            }
        }

    }
}
