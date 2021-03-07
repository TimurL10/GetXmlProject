using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GetXml.Jobs;
using GetXml.Models;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Hangfire.SqlServer;
using Hangfire;
using Microsoft.AspNetCore.Http;

namespace GetXml
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            //var connString = Configuration.GetValue<string>("DbInfo:ConnectionString");
            //GlobalConfiguration.Configuration.UseSqlServerStorage(connString);
            services.AddHangfire(configuration => configuration
            .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
            .UseSimpleAssemblyNameTypeSerializer()
            .UseRecommendedSerializerSettings()
            .UseSqlServerStorage(Configuration.GetConnectionString("HangfireConnection"), new SqlServerStorageOptions
            {
                CommandBatchMaxTimeout = TimeSpan.FromMinutes(5),
                SlidingInvisibilityTimeout = TimeSpan.FromMinutes(5),
                QueuePollInterval = TimeSpan.Zero,
                UseRecommendedIsolationLevel = true,
                DisableGlobalLocks = true
            }));
            services.AddHangfireServer();
            services.AddTransient<IMyJob, MyJob>();
            services.AddTransient<IHLogic, HLogic>();
            services.AddTransient<IDeviceRepository, DeviceRepository>();
            services.AddSingleton<IConfiguration>(Configuration);
            services.AddControllersWithViews();
            services.AddScoped<IHangfireJobScheduler, HangfireJobScheduler>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        [Obsolete]
        public void Configure(
        IApplicationBuilder app,
        IWebHostEnvironment env,
        IServiceProvider serviceProvider,
        IMyJob myJob,
        IRecurringJobManager recurringJobManager,
        IHangfireJobScheduler hangfireJobScheduler
        )
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
            });

            app.UseHangfireDashboard("/hangfire", new DashboardOptions
            {
                Authorization = new[] { new HangfireDashboardAuthorizationFilter() }
            });
            
            app.UseHangfireServer(new BackgroundJobServerOptions
            {
                WorkerCount = 2,
                SchedulePollingInterval = TimeSpan.FromMinutes(2)
            });

            BackgroundJob.Enqueue(() => serviceProvider.GetService<IHangfireJobScheduler>().ScheduleRecurringJobs());
        }
    }
}
