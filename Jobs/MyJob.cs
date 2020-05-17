using Hangfire;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Hangfire.Dashboard;
using Hangfire.Annotations;
using System.Security.Claims;
using System.Net.Http;
using System.Xml.Serialization;
using GetXml.Models;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Threading;

namespace GetXml.Jobs
{
    public interface IMyJob
    {
        Task RunAtTimeOf(DateTime now);
    }

    public class MyJob : IMyJob
    {
        private ILoggerFactory _loggerFactory;
        public IConfiguration configuration;
        private readonly DeviceRepository deviceRepository;
        public MyJob(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _loggerFactory = loggerFactory;
            deviceRepository = new DeviceRepository(configuration);
            loggerFactory.AddFile(Path.Combine(Directory.GetCurrentDirectory(), "logger.txt"));
            deviceRepository = new DeviceRepository(configuration);
        }
        
        public async Task Run(IJobCancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            await RunAtTimeOf(DateTime.Now);
        }

        public async Task RunAtTimeOf(DateTime now)
        {
            GetTerminals();
        }

        public async void GetTerminals()
        {
            var loggerF = _loggerFactory.CreateLogger("FileLogger");
            try
            {
                var client = new HttpClient();
                string GETXML_PATH = "https://api.ar.digital/v4/devices/xml/M9as6DMRRGJqSVxcE9X58TM2nLbmR99w";
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
                                deviceRepository.Update(deviceFromDb);
                            }
                            else
                            {
                                var deviceFromDb = deviceRepository.Get(d.Id);
                                deviceFromDb.SumHours += deviceFromDb.Hours_Offline;
                                deviceFromDb.Hours_Offline = d.Hours_Offline;
                                deviceFromDb.Last_Online = d.Last_Online;
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
    }

    public class HangfireJobScheduler
    {
        [Obsolete]
        public static void ScheduleRecurringJobs()
        {
            RecurringJob.RemoveIfExists(nameof(MyJob));
            RecurringJob.AddOrUpdate<MyJob>(nameof(MyJob),
                job => job.Run(JobCancellationToken.Null),
                Cron.MinuteInterval(5), TimeZoneInfo.Utc);
        }
    }

    public class HangfireDashboardAuthorizationFilter : IDashboardAuthorizationFilter
    {
        public bool Authorize ([NotNull] DashboardContext context)
        {
            var httpcontext = context.GetHttpContext();
            var userRole = httpcontext.User.FindFirst(ClaimTypes.Role)?.Value;
            return true;
        }
    }

   
}
