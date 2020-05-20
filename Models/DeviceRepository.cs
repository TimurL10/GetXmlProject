using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace GetXml.Models
{
    public class DeviceRepository : IDeviceRepository
    {
        private readonly string connectionString;        
        public DeviceRepository(IConfiguration configuration)
        {
            connectionString = configuration.GetValue<string>("DbInfo:ConnectionString");
        }

        internal IDbConnection Connection
        {
            get
            {
                return new SqlConnection(connectionString);
            }
        }
        public List<Device> GetDevices()
        {          
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<Device>("Select terminal.*,address.address From terminal Left Join address On address.name = terminal.name order by SumHours DESC").ToList();
            }
        }

        public Device Get(double id)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<Device>("Select * From terminal Where id = @id", new { Id = id }).FirstOrDefault();
            }
        }

        public void Update(Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Update terminal Set id = @Id, name = @Name, status = @Status, ip = @Ip, last_online = @Last_Online, campaign_name = @Campaign_Name, hours_offline = @Hours_Offline, SumHours = @SumHours Where id = @Id", device);
                dbConnection.Close();
            }
        }
        public void Delete(double id)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("DELETE FROM terminal WHERE Id = @id", new { Id = id });
                dbConnection.Close();
            }
        }

        public void Add(Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Insert Into terminal(id, name, status, ip, last_online, campaign_name, SumHours) Values (@Id, @Name, @Status, @Ip, @Last_Online, @Campaign_Name, @SumHours)", device);
                dbConnection.Close();
            }
        }

        public void AddAddress(Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Insert Into address (name, address) Values (@Name, @Address)", device);
                dbConnection.Close();
            }               
        }

        public List<Device> GetDataForReport()
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                //return dbConnection.Query<Device>("Select t.id,t.name,t.status, t.campaign_name, t.ip, t.last_online,a.address, t.hours_offline,t.SumHours From terminal t Join address a On a.name = t.name order by SumHours DESC").Select(f => new Device(f.Id,f.Name,f.Status,f.Campaign_Name,f.Ip,f.Last_Online,f.Address,f.Hours_Offline,f.SumHours)).ToList();
                return dbConnection.Query("Select t.id,t.name,t.status, t.campaign_name, t.ip, t.last_online,a.address, t.hours_offline,t.SumHours From terminal t Join address a On a.name = t.name order by SumHours DESC").Select(f => new Device(f.Id, f.Name, f.Status, f.Campaign_Name, f.Ip, f.Last_Online, f.Address, f.Hours_Offline, f.SumHours)).ToList();
            }
        }
    }
}
