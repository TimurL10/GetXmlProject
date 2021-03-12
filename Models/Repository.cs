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
using GetXml.Models;
using System.Diagnostics;
using System.Collections;

namespace FarmacyControl
{
    public class Repository : IRepository
    {
        private string _configuration;
        private string _configCloud;

        public Repository(IConfiguration configuration)
        {
            _configuration = configuration.GetValue<string>("DataBaseInfo:ConnectionString");
            _configCloud = configuration.GetValue<string>("DBInfo:ConnectionString");
        }

        internal IDbConnection dbConnection
        {
            get
            {
                return new SqlConnection(_configuration);
            }

        }

        internal IDbConnection dbConnectionCloud
        {
            get
            {
                return new SqlConnection(_configCloud);
            }
        }

        public List<Mrc> GetMrc()
        {
            using (IDbConnection connection = dbConnection)
            {
                return connection.Query<Mrc>("SELECT  *  FROM [References].[ManufactorGoodsPrices]").ToList();
            }
        }

        public void UpdateDb(Mrc mrcs)
        {
            using (IDbConnection connection = dbConnection)
            {
                connection.Execute($"Update [References].[ManufactorGoodsPrices] set NNT = {mrcs.Nnt},Price = {mrcs.Price} where NNT = {mrcs.Nnt}", mrcs);
            }
        }

        public void InsertDb(Mrc mrcs)
        {
            using (IDbConnection connection = dbConnection)
            {
                connection.Execute("Insert into [References].[ManufactorGoodsPrices] (NNT, Price) Values(@Nnt, @Price)", mrcs);
            }
        }

        public List<Market> GetMarkets()
        {
            ArrayList arrayList = new ArrayList();
            //List<Market> markets = new List<Market>();
            //using (IDbConnection connection = dbConnection)
            //{
            //    var sql = "exec [Monitoring].[GetOffStoresFromSite] @deep=1";
            //    int i = 1;
            //    markets = connection.Query<Market>(sql).ToList();

            //}

            List<Market> markets = new List<Market>();
            using (IDbConnection connection = dbConnection)
            {
                string spName = "[Monitoring].[GetOffStoresFromSite]";

                SqlCommand command = new SqlCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.Connection = (SqlConnection)connection;
                command.CommandText = spName;
                SqlParameter parameter = new SqlParameter();
                parameter.ParameterName = "@deep";
                parameter.SqlDbType = SqlDbType.NVarChar;
                parameter.Direction = ParameterDirection.Input;
                parameter.Value = 1;
                command.Parameters.AddWithValue("@deep", 1);
                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            //var parser = reader.GetRowParser<Market>(typeof(Market));
                            //var myObject = parser(reader);
                            //markets.Add(myObject);
                            for (var i = 0; i < 8; i ++)
                            {
                                if (reader[i] == null || reader[i] == DBNull.Value && i == 4)
                                    arrayList.Add(DateTimeOffset.MinValue);
                                else if (reader[i] == null || reader[i] == DBNull.Value)
                                    arrayList.Add("-");
                                else
                                    arrayList.Add(reader[i]);
                            }
                            
                                Market market = new Market((Guid)arrayList[0], (string)arrayList[1], (string)arrayList[2],
                                    (string)arrayList[3], (DateTimeOffset)arrayList[4], (bool)arrayList[5], (bool)arrayList[6], (bool)arrayList[7]);
                                markets.Add(market);
                            
                            arrayList.Clear();
                        }


                    }
                    else
                    {
                        Console.WriteLine("No rows found.");
                    }
                    reader.Close();
                }
            }
            return markets;
        }        

        public void InsertMarkets(Market market)
        {
            using (IDbConnection connection = dbConnectionCloud)
            {
                connection.Execute("Insert into MarketsActivity (StoreId, StoreName, NetName, SoftwareName, StockDate, ActiveFl, reserveFl, StocksFl, TimeStamp, Reason, Status) Values (@StoreId, @StoreName, @NetName, @SoftwareName, @StockDate, @ActiveFl, @reserveFl, @StocksFl, @TimeStamp, @Reason, @Status) ", market);
            }
        }
        public List<Market> GetSavedMarkets()
        {
            using (IDbConnection connection = dbConnectionCloud)
            {
                return connection.Query<Market>("Select * from MarketsActivity").ToList();
            }
        }
        public void UpdateMarkets(Market market)
        {
            using (IDbConnection connection = dbConnectionCloud)
            {
                connection.Execute($"Update MarketsActivity set Status = {market.Status}, Reason={market.Reason}, TimeStamp = {market.TimeStamp}", market);
            }
        }
        public void DeleteActiveMarket(Market market)
        {
            using (IDbConnection connection = dbConnectionCloud)
            {
                connection.Execute($"Delete from MarketsActivity Where StoreId = {market.StoreId}", market);
            }
        }
    }        
}

