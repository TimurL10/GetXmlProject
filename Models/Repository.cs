using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Reflection.Metadata.Ecma335;
using Microsoft.Data.SqlClient;
using FarmacyControl.Models;

namespace FarmacyControl
{
    public class Repository : IRepository
    {
        private string _configuration;

        public Repository(IConfiguration configuration)
        {
            _configuration = configuration.GetValue<string>("DataBaseInfo:ConnectionString");
        }

        internal IDbConnection dbConnection
        {
            get
            {
                return new SqlConnection(_configuration);
            }

        }   

        public List<Mrc> GetMrc()
        {
            using (IDbConnection connection = dbConnection)
            {
                return connection.Query<Mrc>("SELECT  *  FROM [References].[ManufactorGoodsPrices]").ToList();
            }
        }

        public void WriteToDb(Mrc mrcs)
        {
            using (IDbConnection connection = dbConnection)
            {
                connection.Execute("Insert into [References].[ManufactorGoodsPrices] (NNT, Price) Values(@Nnt, @Price)", mrcs);
            }
        }
       

    }
}
