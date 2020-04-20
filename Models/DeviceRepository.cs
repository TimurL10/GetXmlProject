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
                return new NpgsqlConnection(connectionString);
            }
        }
        public List<Device> GetDevices()
        {          
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<Device>("Select terminal.*,address.address From terminal Left Join address On address.name = terminal.name").ToList();
            }
        }

        public Device Get(double id)
        {
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                return dbConnection.Query<Device>("Select * From terminal Where id = @id", new { Id = id }).FirstOrDefault();
            }
        }

        public void Update(Device device)
        {
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Update terminal Set id = @Id, name = @Name, status = @Status, ip = @Ip, last_online = @Last_Online, campaign_name = @Campaign_Name, hours_offline = @Hours_Offline Where id = @Id", device);
            }
        }
        public void Delete(double id)
        {
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("DELETE FROM terminal WHERE Id = @id", new { Id = id });
            }
        }

        public void Add(Device device)
        {
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Insert Into terminal(id, name, status, ip, last_online, campaign_name) Values (@Id, @Name, @Status, @Ip, @Last_Online, @Campaign_Name)", device);
            }
        }

        public void AddAddress(Device device)
        {
            using (var dbConnection = new NpgsqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute("Insert Into address (name, address) Values (@Name, @Address)", device);
            }               
        }
    }
}
