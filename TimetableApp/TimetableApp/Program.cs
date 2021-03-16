using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


/*
using Excel = Microsoft.Office.Interop.Excel;
*/

namespace TimetableApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelFile file = new ExcelFile("Auditory_types.xl sx");
      
            if (file.exists())
            {

                /*внутри ExcelFile проверка file методами классов ...Check, которые используют интерфейс IError...      */
                //проверка содержимого файла (дубликаты, пустые строки, некорректные значения и т.д.)

                // что делать если бд пуста/заполнена/связные таблицы ...
                // загрузка
            }

            // DBConnectionTest.Test();

            
            Console.ReadKey();
        }
    }
}