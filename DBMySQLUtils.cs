using System;
using MySql.Data.MySqlClient;

namespace Baza
{
    class DBMySQLUtils
    {
        public static MySqlConnection GetDBConnection(string host, int port, string database, string user, string password)
        {
            String connString = "Server=" + host + ";database=" + database + ";port=" + port.ToString() + ";user=" + user + ";password=" + password + ";";
            MySqlConnection conn = new MySqlConnection(connString);
            return conn;
        }
    }
}
