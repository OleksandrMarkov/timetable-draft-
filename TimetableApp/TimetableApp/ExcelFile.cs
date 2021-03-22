using System;
using System.Collections;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

using MySql.Data.MySqlClient;


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

        /*
        public ArrayList f()
        {
            ArrayList recordsInExcelFile = new ArrayList();
            recordsInExcelFile.Add(1);
            recordsInExcelFile.Add(1);
            recordsInExcelFile.Add(1);
            recordsInExcelFile.Add(1);
            recordsInExcelFile.Add(1);
            return recordsInExcelFile;
        }*/

        // завантаження до бази даних
        public void load()
        {
            switch (FileName)
            {
                case "TypesOfAuditories.xlsx":
                    ArrayList recordsInExcelFile = new ArrayList();

                    ArrayList rowsWithMissingValues = new ArrayList();
                    Dictionary<int, string> duplicates = new Dictionary<int, string>();
                    columnForReading = 'A';

                    try
                    {
                        open();
                        for (int row = 1, column = getColumnNumber(columnForReading); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            var cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (recordsInExcelFile.Contains(cellContent) == true)
                            {
                                duplicates.Add(row, cellContent);
                            }
                            else
                            {
                                recordsInExcelFile.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
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

                    if (rowsWithMissingValues.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
                        foreach (int row in rowsWithMissingValues)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (duplicates.Count != 0)
                    {
                        Console.WriteLine("Есть дубликаты:");
                        foreach (KeyValuePair<int, string> duplicate in duplicates)
                        {
                            Console.WriteLine("В строке номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                    }

                    if (rowsWithMissingValues.Count == 0 && duplicates.Count == 0)
                    {
                        try
                        {
                            ArrayList recordsInDB = new ArrayList();
                            MySqlConnection connection = DBUtils.GetDBConnection();
                            MySqlCommand mySqlCommand;
                            MySqlDataReader dataReader;



                            // Отримання даних з БД та порівняння з даними з Excel-файлу
                            // Якщо вони співпадають - немає сенсу для перезапису, інакше дані в БД перезаписуються

                            const string selectAuditoryTypes = "SELECT auditory_type_name FROM auditory_type";
                            connection.Open();
                            mySqlCommand = new MySqlCommand (selectAuditoryTypes, connection);
                            dataReader = mySqlCommand.ExecuteReader();

                            while (dataReader.Read())
                            {
                                recordsInDB.Add(dataReader[0].ToString());
                            }

                            connection.Close();

                            bool noSenseToReload = true;
                            foreach (string record in recordsInExcelFile)
                            {
                                if (!recordsInDB.Contains(record))
                                {
                                    noSenseToReload = false;       
                                    break;
                                }
                            }

/*
                            Console.WriteLine("Excel:");
                            foreach (string record in recordsInExcelFile)
                            {
                                Console.WriteLine(record);
                            }
                            Console.WriteLine();
                            Console.WriteLine("DB:");
                            foreach (string record in recordsInDB)
                            {
                                Console.WriteLine(record);
                            }
*/


                            if (noSenseToReload == false)
                            {
                                Console.WriteLine("Є що змінювати");
                                try
                                {
                                    /* // очищення таблиці в БД
                                     connection.Open();
                                     const string truncateAuditoryTypes = "TRUNCATE TABLE auditory_type";
                                     mySqlCommand = new MySqlCommand(truncateAuditoryTypes, connection);
                                     mySqlCommand.ExecuteNonQuery();
                                     //Console.WriteLine("Очищено");
                                     connection.Close();
                                     */

                                    // перезапис таблиці в БД
                                    const string insertAuditoryTypes = "INSERT INTO auditory_type (auditory_type_name) VALUES (@TYPE)";

                                    connection.Open();
                                    foreach (string record in recordsInExcelFile)
                                    {
                                        mySqlCommand = new MySqlCommand(insertAuditoryTypes, connection);
                                        mySqlCommand.Parameters.AddWithValue("@TYPE", record);
                                        mySqlCommand.ExecuteNonQuery();
                                        Console.WriteLine(record);
                                    }

                                    connection.Close();
                                }
                                catch
                                {
                                    Console.WriteLine("Не вдалося виконати перезапис в базі даних, оскільки є залежність між даними!");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
                        }
                    }
                    break;
                case "Disciplines.xlsx":
                    break;
                case "Faculties.xlsx":
                    break;
                case "Departments.xlsx":
                    break;
                case "Teachers.xlsx":
                    break;
                case "Auditories.xls":
                    break;
                case "StudyGroups.xlsx":
                    break;
                default:
                    throw new Exception("Невідомий файл");
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

