using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;

namespace TimetableInConsole.DBConnection
{

	public class DBConnectionTest
	{
public static void Test()
        {
            Console.WriteLine("Getting connection...");

            MySqlConnection connection = DBUtils.GetDBConnection();

            try
            {
                Console.WriteLine("Opening connection...");
                connection.Open();
                Console.WriteLine("Connection Successful!");
                connection.Close();

            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }

            finally
            {
                Console.Read();
            }
        }
	}
}
