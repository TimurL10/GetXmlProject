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
        void RunTwo(IJobCancellationToken token);
    }

    public class MyJob : IMyJob
    {        
        private IHLogic _hLogic;
        public MyJob(IHLogic hLogic)
        {           
            _hLogic = hLogic;
        }        
       
        public void RunTwo(IJobCancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            RunAtTimeOfActivity(DateTime.Now);
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

            //RecurringJob.RemoveIfExists("hour terminal activity updating");
            //RecurringJob.AddOrUpdate<MyJob>("hour terminal activity updating",
            //    job => job.RunTwo(JobCancellationToken.Null),
            //    "0 0 ? * * *", TimeZoneInfo.Utc); 
            
            RecurringJob.RemoveIfExists("hour terminal activity updating");
            RecurringJob.AddOrUpdate<MyJob>("hour terminal activity updating",
                job => job.RunTwo(JobCancellationToken.Null), Cron.MinuteInterval(3),              
                TimeZoneInfo.Utc);
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
