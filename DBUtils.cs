using MySql.Data.MySqlClient;

namespace Baza
{
    class DBUtils
    {
        public static MySqlConnection GetDBConnection()
        {
            string host = "localhost";
            int port = 3306;
            string database = "baza_db";
            string user = "root";
            string password = "root";
            return DBMySQLUtils.GetDBConnection(host, port, database, user, password);
        }
    }
}
