using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO; 

namespace TimetableApp
{
    /*
     Проверка существования Excel-файла
     в директории
     */
    class FileExistence
    {
        string fileName;
        const string dirName = @"E:\BACHELORS WORK\TIMETABLE\data";

        public FileExistence(string fileName)
        {
            this.fileName = fileName;
        }

        public bool Exists()
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
        }
    }
}

