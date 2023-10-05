using Dapper;
using Northern_Ireland_Shipments.Infrastructure.FileBroker;
using Northern_Ireland_Shipments.Interfaces;
using Northern_Ireland_Shipments.Models.Queries;
using Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig;
using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace Northern_Ireland_Shipments.Application.Db
{
    public class RpDataExtractToList : IRpDataExtractToList
    {
        private static RpDataExtractToList rpDataExtractToList;
        private static string rpDbProdConn, query;

        public RpDataExtractToList()
        {
            rpDbProdConn = RpOracleDbConnectionProd.Read();
            query = SqlQueryRead.Instance.RpDbQueryRead();
        }

        public static RpDataExtractToList Instance
        {
            get
            {
                if (rpDataExtractToList == null)
                    rpDataExtractToList = new RpDataExtractToList();
                return rpDataExtractToList;
            }
        }

        protected IDbConnection GetConnection
        {
            get
            {
                var oracleConnection = new OracleConnection(rpDbProdConn);
                oracleConnection.Open();
                return oracleConnection;
            }
        }

        public List<RpDbQueryModel> GetList()
        {
            using (var dbConnection = GetConnection)
            {
                return dbConnection.Query<RpDbQueryModel>(query).ToList();
            }
        }
    }
}
