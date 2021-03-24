using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;

namespace TimetableInConsole.DBConnection
{

	public class DBUtils
	{
		public static MySqlConnection GetDBConnection()
        {
            string host = "127.0.0.1";
            int port = 3306;
            string db = "app_db";
            string username = "root";
            string password = "pbx93fq26";

            return DBMySQLUtils.GetDBConnection(host, port, db, username, password);
        }
	}
}
