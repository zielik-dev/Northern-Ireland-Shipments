using Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig;

namespace Northern_Ireland_Shipments.Infrastructure.FileBroker
{
    public class SqlQueryRead : ConnectionStrings
    {
        private static SqlQueryRead sqlQueryRead;
        private readonly string sqlQueriesDir;

        public SqlQueryRead()
        {
            sqlQueriesDir = SqlQueriesDir.Read();
        }

        public static SqlQueryRead Instance
        {
            get
            {
                if (sqlQueryRead == null)
                    sqlQueryRead = new SqlQueryRead();
                return sqlQueryRead;
            }
        }

        public string RpDbQueryRead()
        {
            string fullPath = Path.Combine(sqlQueriesDir, queryFile);
            
            var dir = File.ReadAllText(fullPath);

            return dir;
        }
    }
}