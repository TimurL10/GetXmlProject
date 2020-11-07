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
using GetXml.Controllers;

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
            //_homeController.Index();
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
