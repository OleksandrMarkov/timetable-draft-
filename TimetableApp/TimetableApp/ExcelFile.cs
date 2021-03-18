﻿using System;
using System.Collections;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace TimetableApp
{

    class ExcelFile
    {
        private string fileName;

        private string wrongFileName; // для повідомлень про помилку

        private const string dirName = @"E:\BACHELORS WORK\TIMETABLE\data"; // директорія, де зберігається Excel-файл

        private Excel.Application app;
        private Excel.Workbook workbook; // книга
        private Excel.Worksheet worksheet; // лист 
        private Excel.Range range; // діапазон

        protected int rowsCount; // к-ть рядків файлу з даними
        protected string cellContent; // дані в чарунці


        protected char columnForReading; // стовпець, що розглядається

        // private int columnForLoading; // стовпець для завантаження (А - 1, В - 2, С - 3 і т.д.)



        public string FileName
        {
            get
            {
                return fileName;
            }

            set
            {
                fileName = null; // якщо ім'я файлу буде некоректним в змінній "fileName" залишиться null


                // Перевірка наявності заборонених символів в іменах Excel-файлів
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
                // Перевірка розширення (".xlsx" або ".xls")
                if (containsForbiddenSymbols == false && (value.EndsWith(".xlsx") || value.EndsWith(".xls")))
                {
                    fileName = value;
                    //Console.WriteLine("good name " + fileName);
                }
                else
                {
                    wrongFileName = value;
                    Console.WriteLine("Некоректне ім'я або розширення файлу \"" + wrongFileName +
                        "\". Ім'я файлу не повинно містити наступних символів: " + forbiddenSymbols +
                        "\nРозширення файлу повинно бути \".xls\" або \".xlsx\"\n");

                    // Environment.Exit(0);
                }
            }
        }

        // Повне ім'я файлу
        public string FullPathToFile
        {
            get
            {
                return String.Concat(dirName, @"\", fileName);
            }
        }

        public ExcelFile(string fileName)
        {
            FileName = fileName;

            try
            {
                open();

              /* перенесено в open()
               *range = worksheet.UsedRange;
                rowsCount = range.Rows.Count;*/
                close();
                //Console.WriteLine(rowsCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Не вдалось отримати дані файлу \"" + wrongFileName +
                    "\". Перевірте правильність імені файлу та повторіть спробу.");
            }
        }

        //відкриття файлу
        public void open()
        {
            if (fileName == null)
            {
                Console.WriteLine("Помилка відкриття файлу \"" + wrongFileName +
                    "\". Перевірте правильність імені файлу та повторіть спробу.");
            }
            else
            {
                app = new Excel.Application();

                int readingMode = 0; // режим читання файлу
                int sheetNumber = 1; // інформація в файлах знаходиться на першому листі

                try
                {
                    workbook = app.Workbooks.Open(FullPathToFile, readingMode, true); // відкриття книги 
                    worksheet = (Excel.Worksheet)workbook.Worksheets[sheetNumber]; // відкриття листа даних
                    //worksheet = app.Worksheets["Лист 1"] as Excel.Worksheet;

                    range = worksheet.UsedRange;
                    rowsCount = range.Rows.Count;

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Помилка відкриття файлу \"" + wrongFileName +
                        "\". Перевірте правильність імені файлу та повторіть спробу.");
                }
            }
        }

        // Перевірка існування файлу в директорії
        public bool exists()
        {
            if (File.Exists(FullPathToFile))
            {
                //Console.WriteLine("OK");
                return true;
            }
            else
            {
                Console.WriteLine("Файл не знайдено: " + "\"" + wrongFileName + "\"");
                return false;
            }
        }

        //закриття файлу
        public void close()
        {
            try
            {
                workbook.Close(false);
                app.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Помилка при закритті файлу: \"" + wrongFileName + "\".");
            }
        }

        protected int getColumnNumber(char column)
        {
            if ((int)column >= 'A' && (int)column <= 'Z')
            {
                return (int)column - 64;
            }
            else
            {
                throw new Exception("Некоректне ім'я стовпця, неможливо зчитати дані з файлу " + FileName);
            }

        }

        // перевірка наявності дублікатів
        public bool containsDuplicates()
        {
            if (FileName == "TypesOfAudiences.xlsx")
            {
                ArrayList recordsList = new ArrayList();
                Dictionary<int, string> duplicates = new Dictionary<int, string>();
                columnForReading = 'A';
                // Console.WriteLine(rowsCount);

                try
                {
                    open();
                    for (int row = 1, column = getColumnNumber(columnForReading); row <= rowsCount; row++)
                    {
                        //Console.WriteLine(row);
                        cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                        //Console.WriteLine("cellContent: " + cellContent);
                        if (recordsList.Contains(cellContent) == true)
                        {
                            duplicates.Add(row, cellContent);
                        }
                        else
                        {
                            recordsList.Add(cellContent);
                        }
                    }
                    close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                }

                if (duplicates.Count == 0)
                {
                    //Console.WriteLine("duplicates.Count = " + duplicates.Count);
                    Console.WriteLine("Нет дубликатов");
                    return false;
                }
                else
                {
                    Console.WriteLine("Есть дубликаты:");
                    foreach (KeyValuePair<int, string> duplicate in duplicates)
                    {
                        Console.WriteLine("В строке номер " + duplicate.Key + ": " + duplicate.Value);
                    }
                    return true;
                }
            }
            else
            {
                Console.WriteLine("другой файл...");
                return false;
            }
        }

        // перевірка наявності прогалин
        public bool containsMissingValues()
        {
            if (FileName == "TypesOfAudiences.xlsx")
            {
                ArrayList rowsWithMissingValues = new ArrayList();
                columnForReading = 'A';

                try
                {
                    open();
                    for (int row = 1, column = getColumnNumber(columnForReading); row <= rowsCount; row++)
                    {
                        //Console.WriteLine(row);

                        var cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                        //Console.WriteLine("cellContent: " + cellContent);
                        if (string.IsNullOrEmpty(cellContent))
                        {
                            rowsWithMissingValues.Add(row);
                        }
                    }
                    close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                }

                if (rowsWithMissingValues.Count == 0)
                {
                    Console.WriteLine("Немає пропусків в файлі " + FileName);
                    return false;
                }
                else
                {
                    Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
                    foreach (int row in rowsWithMissingValues)
                    {
                        Console.Write(row + "\t");
                    }
                    Console.WriteLine();
                    return true;     
                }
            }
            else
            {
                Console.WriteLine("другой файл...");
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
