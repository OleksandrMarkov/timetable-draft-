using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimetableApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // DBConnectionTest.Test();

            FileExistence file = new FileExistence("Auditory_types.xlsx");

            if (file.Exists())
            {
                //проверка содержимого файла (дубликаты, пустые строки, некорректные значения и т.д.)

                // что делать если бд пуста/заполнена/связные таблицы ...

                // загрузка

            }

            //Console.WriteLine(String.Concat("1", "2343", "asda"));

            Console.ReadKey();
        }
    }
}