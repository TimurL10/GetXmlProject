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

        public void PostActivity(double id, DateTime date, int active)
        {
            using (var dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                dbConnection.Execute($"Insert Into terminal_activity(t_id, active_date, active) Values ({id},convert(varchar, DATEADD(HOUR, +3, GETUTCDATE())),{active})");
                dbConnection.Close();
            }
        }
    }
}
