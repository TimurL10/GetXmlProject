using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GetXml.Models;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Http;
using GetXml.HLogicFolder;
using FarmacyControl;

namespace GetXml
{
    public class Startup
    {
        public IConfiguration Configuration { get; }

        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }        

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {            
            services.AddTransient<IHLogic, HLogic>();
            services.AddTransient<IDeviceRepository, DeviceRepository>();
            services.AddSingleton<IConfiguration>(Configuration);
            services.AddControllersWithViews();
            services.AddScoped<IMarketsLogic, MarketsLogic>();
            services.AddScoped<IRepository, Repository>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        [Obsolete]
        public void Configure(
        IApplicationBuilder app,
        IWebHostEnvironment env,
        IServiceProvider serviceProvider              
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

            //app.UseHangfireDashboard("/hangfire", new DashboardOptions
            //{
            //    Authorization = new[] { new HangfireDashboardAuthorizationFilter() }
            //});
            
        }
    }
}
