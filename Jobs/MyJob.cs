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
using System.Diagnostics;

namespace GetXml.Jobs
{
    public interface IMyJob
    {
        void RunAtTimeOf(DateTime now);
        void RunAtTimeOfActivity(DateTime now);
        void RunTwo(IJobCancellationToken token);
    }

    public class MyJob : IMyJob
    {        
        private IHLogic _hLogic;
        public MyJob(IHLogic hLogic)
        {           
            _hLogic = hLogic;
        }
        
        public void Run(IJobCancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            RunAtTimeOf(DateTime.Now);
        }
        public void RunTwo(IJobCancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            RunAtTimeOfActivity(DateTime.Now);
        }

        public void RunAtTimeOf(DateTime now)
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
        
        public void RunAtTimeOfActivity(DateTime now)
        {           
            Task task1 = new Task(() => _hLogic.GetXmlData());
            task1.Start();
            task1.Wait();            

            Task task3 = new Task(() => _hLogic.SaveActivity());
            task3.Start();
            task3.Wait();
        }
    }
    public interface IHangfireJobScheduler
    {
        void ScheduleRecurringJobs();
    }

    public class HangfireJobScheduler : IHangfireJobScheduler
    {
        [Obsolete]
        public void ScheduleRecurringJobs()
        {
            RecurringJob.RemoveIfExists("25 min updating");
            RecurringJob.AddOrUpdate<MyJob>("25 min updating",
                job => job.Run(JobCancellationToken.Null),
                Cron.MinuteInterval(25),TimeZoneInfo.Utc);

            RecurringJob.RemoveIfExists("hour updating");
            RecurringJob.AddOrUpdate<MyJob>("hour updating",
                job => job.RunTwo(JobCancellationToken.Null),
                "0 0 ? * * *", TimeZoneInfo.Utc);
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
