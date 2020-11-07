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
        private IHLogic _hLogic;
        private readonly DeviceRepository deviceRepository;
        public MyJob(ILoggerFactory loggerFactory, IConfiguration configuration, IHLogic hLogic)
        {
            _loggerFactory = loggerFactory;
            deviceRepository = new DeviceRepository(configuration);
            loggerFactory.AddFile(Path.Combine(Directory.GetCurrentDirectory(), "logger.txt"));
            deviceRepository = new DeviceRepository(configuration);
            _hLogic = hLogic;
        }
        
        public async Task Run(IJobCancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            await RunAtTimeOf(DateTime.Now);
        }

        public async Task RunAtTimeOf(DateTime now)
        {
            Task task1 = new Task(() => _hLogic.GetXmlData());
            task1.Start();
            task1.Wait();

            Task task2 = new Task(() => _hLogic.FilterDevices());
            task2.Start();
            task2.Wait();

            Task task3 = new Task(() => _hLogic.getHoursOffline());
            task3.Start();
            task3.Wait();
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
                Cron.MinuteInterval(25),TimeZoneInfo.Utc);
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
