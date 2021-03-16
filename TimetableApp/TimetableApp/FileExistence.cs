using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace TimetableApp
{
    /*
     Проверка существования Excel-файла в директории
     */

    class FileExistence
    {
        private string fileName;
        private const string dirName = @"E:\BACHELORS WORK\TIMETABLE\data";

        public string FileName
        {
            get
            {
                return fileName;
            }

            set
            {
                bool containsForbiddenSymbols = false;

                string forbiddenSymbols = @"\" + "/:*?\"<>|";

                foreach (char ch in forbiddenSymbols)
                {
                    if (value.IndexOf(ch) != -1)
                    {
                        containsForbiddenSymbols = true;
                        break;
                    }
                }

                if (containsForbiddenSymbols == false && (value.EndsWith(".xlsx") || value.EndsWith(".xls")))
                {
                    fileName = value;
                   // Console.WriteLine("good name " + fileName);
                }
                else
                {
                    Console.WriteLine("Некорректное имя файла: " + value + ". Имя файла не должно содержать следующих знаков: " + forbiddenSymbols);
                }
            }
        }

        public FileExistence(string fileName)
        {
            FileName = fileName;
        }

        public bool exists()
        {
            string full_path_to_file = String.Concat(dirName, @"\", fileName);

            if (File.Exists(full_path_to_file))
            {
                //Console.WriteLine("OK");
                return true;
            }
            else
            {
                Console.WriteLine("Файл не найден: " + "\"" + full_path_to_file + "\"");
                return false;
            }
        }
    }
}



/*
{
    string full_path_to_file = String.Concat(dirName, @"\", fileName);

    if (Directory.Exists(dirName))
    {
        string[] files = Directory.GetFiles(dirName);

        if (files.Contains(full_path_to_file))
        {
            return true;
        }
        else
        {
            Console.WriteLine("Файл не обнаружен: " +  "\"" + full_path_to_file + "\"");
            return false;
        }              
    }
    else
    {
        Console.WriteLine("");
        Console.WriteLine("Пути \"" + dirName + "\" не существует!");
        return false;
    }
}*/

