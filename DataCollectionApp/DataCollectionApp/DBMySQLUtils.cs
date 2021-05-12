using System;
using MySql.Data.MySqlClient;

namespace DataCollectionApp
{
	public static class DBMySQLUtils
	{
		public static MySqlConnection
            GetDBConnection(string host, int port, string db, string username, string password)
        {
            String ConnectionString = "Server=" + host + ";Database=" + db
                + ";port=" + port + ";User Id=" + username + ";password=" + password;
            MySqlConnection connection = new MySqlConnection(ConnectionString);
            return connection;
        }
	}
}