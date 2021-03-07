using Hangfire;
using Hangfire.Annotations;
using Hangfire.Dashboard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

namespace GetXml.Jobs
{
    public interface ICheckActivity
    {
        Task RunAtTimeOf(DateTime now);
        Task Run(IJobCancellationToken token);
    }

    public class CheckActivity : ICheckActivity
    {        
        private IHLogic _hLogic;
        public CheckActivity(IHLogic hLogic)
        {
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

            Task task3 = new Task(() => _hLogic.SaveActivity());
            task3.Start();
            task3.Wait();
        }
    }       
}

