using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.HomeLogic
{
    public class HomeLogic
    {
        //public Xml Devices;
        //private ILoggerFactory _loggerFactory;
        //private readonly DeviceRepository deviceRepository;
        //public static string displayName = "(GMT+03:00) Russia Time Zone 2";
        //public static string standardName = "Russia Time Zone 2";
        //public static TimeSpan offset = new TimeSpan(03, 00, 00);
        //public TimeZoneInfo moscow = TimeZoneInfo.CreateCustomTimeZone(standardName, offset, displayName, standardName);
        //public List<string> excelDataAddress = new List<string>();
        //public List<Tuple<double, string>> excelDataNotes = new List<Tuple<double, string>>();

        //public HomeController(ILoggerFactory loggerFactory, IConfiguration configuration)
        //{
        //    _loggerFactory = loggerFactory;
        //    deviceRepository = new DeviceRepository(configuration);
        //    loggerFactory.AddFile(Path.Combine(Directory.GetCurrentDirectory(), "logger.txt"));
        //    deviceRepository = new DeviceRepository(configuration);
        //}

        //public double getHoursOffline(Device device)
        //{
        //    double h_new = 0;
        //    try
        //    {
        //        var Hours_Offline = (DateTime.UtcNow - device.Last_Online).TotalHours;
        //        var ts_new = TimeSpan.FromHours(Hours_Offline);
        //        h_new = System.Math.Floor(ts_new.TotalHours);
        //        device.Hours_Offline = h_new;

        //        if (h_new < 0)
        //            h_new = 0;
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine(e.Message);
        //    }

        //    return h_new;
        //}

        //public void getHoursOffline()
        //{
        //    double h_new = 0;
        //    try
        //    {
        //        foreach (var d in Devices.Devices.ToList())
        //        {
        //            AddNewDevice(d);
        //            //var Hours_Offline = (DateTime.UtcNow - d.Last_Online).TotalHours;
        //            //var ts_new = TimeSpan.FromHours(Hours_Offline);
        //            //h_new = System.Math.Floor(ts_new.TotalHours);                    
        //            var deviceFromDb = deviceRepository.Get(d.Id);
        //            if (d.Hours_Offline == deviceFromDb.Hours_Offline)
        //                ChangeTerminalData(d);
        //            else
        //                OfflineHoursCount(d);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine(e.Message);
        //    }
        //}





    }
}
