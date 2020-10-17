using Dapper;
using Microsoft.CodeAnalysis.CSharp.Syntax;
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
                return dbConnection.Query<Device>("Select * From terminal order by hours_offline DESC").ToList();
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

        public void UpdateAddress(Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Update terminal Set name = @Name, address = @Address Where name = @Name", device);
                dbConnection.Close();
            }               
        }

        public void AddAddress(Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                //dbConnection.Execute("Insert Into address (name, address) Values (@Name, @Address)", device);
                dbConnection.Execute("Update terminal Set address = @Address Where name = @Name", device);
            }
        }

        public Device GetAddress(string name)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<Device>("Select * from terminal where name = @Name", new {Name = name }).FirstOrDefault();
            }
        }

        public void UpdateDevice (Device device)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Update terminal Set teamviewer = @TeamViewer, address = @Address, note = @Note Where name = @Name", device);
                dbConnection.Close();
            }
        }
    }
}
