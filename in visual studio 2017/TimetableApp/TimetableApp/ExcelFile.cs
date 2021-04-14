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


        protected char columnForLoading; // стовпець, що розглядається

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
            ArrayList recordsInExcelFileAudTypes = new ArrayList();
            recordsInExcelFileAudTypes.Add(1);
            recordsInExcelFileAudTypes.Add(1);
            recordsInExcelFileAudTypes.Add(1);
            recordsInExcelFileAudTypes.Add(1);
            recordsInExcelFileAudTypes.Add(1);
            return recordsInExcelFileAudTypes;
        }*/

        // завантаження до бази даних
        public void load()
        {
            switch (FileName)
            {
                case "TypesOfAuditories.xlsx":
                    ArrayList recordsInExcelFileAudTypes = new ArrayList();
                    ArrayList missingValuesInExcelFileAudTypes = new ArrayList();
                    Dictionary<int, string> duplicatesInExcelFileAudTypes = new Dictionary<int, string>();
                    columnForLoading = 'A';

                    try
                    {
                        open();
                        for (int row = 1, column = getColumnNumber(columnForLoading); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            var cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (recordsInExcelFileAudTypes.Contains(cellContent) == true)
                            {
                                duplicatesInExcelFileAudTypes.Add(row, cellContent);
                            }
                            else
                            {
                                recordsInExcelFileAudTypes.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                                missingValuesInExcelFileAudTypes.Add(row);
                            }
                        }
                        close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }

                    if (missingValuesInExcelFileAudTypes.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
                        foreach (int row in missingValuesInExcelFileAudTypes)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (duplicatesInExcelFileAudTypes.Count != 0)
                    {
                        Console.WriteLine("Є дублікати:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesInExcelFileAudTypes)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                    }

                    if (missingValuesInExcelFileAudTypes.Count == 0 && duplicatesInExcelFileAudTypes.Count == 0)
                    {
                        try
                        {
                            ArrayList AudTypesInDB = new ArrayList();
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
                                AudTypesInDB.Add(dataReader[0].ToString());
                            }

                            connection.Close();

                            bool noSenseToReload = true;
                            foreach (string record in recordsInExcelFileAudTypes)
                            {
                                if (!AudTypesInDB.Contains(record))
                                {
                                    noSenseToReload = false;       
                                    break;
                                }
                            }

/*
                            Console.WriteLine("Excel:");
                            foreach (string record in recordsInExcelFileAudTypes)
                            {
                                Console.WriteLine(record);
                            }
                            Console.WriteLine();
                            Console.WriteLine("DB:");
                            foreach (string record in AudTypesInDB)
                            {
                                Console.WriteLine(record);
                            }
*/


                            if (noSenseToReload == false)
                            {
                                //Console.WriteLine("Є що змінювати");
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
                                    foreach (string record in recordsInExcelFileAudTypes)
                                    {
                                        mySqlCommand = new MySqlCommand(insertAuditoryTypes, connection);
                                        mySqlCommand.Parameters.AddWithValue("@TYPE", record);
                                        mySqlCommand.ExecuteNonQuery();
                                        //Console.WriteLine(record);
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

                    ArrayList recordsInExcelFileDisciplines = new ArrayList();
                    ArrayList missingValuesInExcelFileDisciplines = new ArrayList();
                    Dictionary<int, string> duplicatesInExcelFileDisciplines = new Dictionary<int, string>();
                    columnForLoading = 'G';
                    try
                    {
                        open();
                        for (int row = 2, column = getColumnNumber(columnForLoading); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (recordsInExcelFileDisciplines.Contains(cellContent) == true)
                            {
                                duplicatesInExcelFileDisciplines.Add(row, cellContent);
                            }
                            else
                            {
                                recordsInExcelFileDisciplines.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                                missingValuesInExcelFileDisciplines.Add(row);
                            }
                        }
                        close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }

                    if (missingValuesInExcelFileDisciplines.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
                        foreach (int row in missingValuesInExcelFileDisciplines)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (duplicatesInExcelFileDisciplines.Count != 0)
                    {
                        Console.WriteLine("Є дублікати:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesInExcelFileDisciplines)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                    }

                    //if (duplicatesInExcelFileDisciplines.Count == 0 && missingValuesInExcelFileDisciplines.Count == 0)
                    //{
                    try
                    {
                        ArrayList DisciplinesInDB = new ArrayList();
                        MySqlConnection connection = DBUtils.GetDBConnection();
                        MySqlCommand mySqlCommand;
                        MySqlDataReader dataReader;

                        // Отримання даних з БД та порівняння з даними з Excel-файлу
                        // Якщо вони співпадають - немає сенсу для перезапису, інакше дані в БД перезаписуються

                        const string selectDisciplines = "SELECT full_name FROM discipline";
                        connection.Open();
                        mySqlCommand = new MySqlCommand(selectDisciplines, connection);
                        dataReader = mySqlCommand.ExecuteReader();

                        while (dataReader.Read())
                        {
                            DisciplinesInDB.Add(dataReader[0].ToString());
                        }
                        connection.Close();

                        bool noSenseToReload = true;
                        foreach (string record in recordsInExcelFileDisciplines)
                        {
                            if (!DisciplinesInDB.Contains(record))
                            {
                                noSenseToReload = false;
                                break;
                            }
                        }

                        if (noSenseToReload == false)
                        {
                            //Console.WriteLine("Є що змінювати");
                            try
                            {
                                /* // очищення таблиці в БД
                                 connection.Open();
                                 const string truncateDisciplines = "TRUNCATE TABLE discipline";
                                 mySqlCommand = new MySqlCommand(truncateDisciplines, connection);
                                 mySqlCommand.ExecuteNonQuery();
                                 //Console.WriteLine("Очищено");
                                 connection.Close();
                                 */

                                // перезапис таблиці в БД
                                const string insertDisciplines = "INSERT INTO discipline (full_name) VALUES (@FULL_NAME)";

                                connection.Open();
                                foreach (string record in recordsInExcelFileDisciplines)
                                {
                                    mySqlCommand = new MySqlCommand(insertDisciplines, connection);
                                    mySqlCommand.Parameters.AddWithValue("@FULL_NAME", record);
                                    mySqlCommand.ExecuteNonQuery();
                                    //Console.WriteLine(record);
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
                    //}

                    break;

                case "Faculties.xlsx":
                    ArrayList namesInExcelFileFaculties = new ArrayList();
                    ArrayList codesInExcelFileFaculties = new ArrayList();

                    ArrayList missingValuesOfNamesInExcelFileFaculties = new ArrayList();
                    ArrayList missingValuesOfCodesInExcelFileFaculties = new ArrayList();

                    Dictionary<int, string> duplicatesOfNamesInExcelFileFaculties = new Dictionary<int, string>();
                    Dictionary<int, string> duplicatesOfCodesInExcelFileFaculties = new Dictionary<int, string>();

                    char columnForLoadingCode = 'D';
                    char columnForLoadingName = 'B';

                    try
                    {
                        open();
                        for (int row = 2, column = getColumnNumber(columnForLoadingName); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (namesInExcelFileFaculties.Contains(cellContent) == true)
                            {
                                duplicatesOfNamesInExcelFileFaculties.Add(row, cellContent);
                            }
                            else
                            {
                                namesInExcelFileFaculties.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                                missingValuesOfNamesInExcelFileFaculties.Add(row);
                            }
                        }

                        for (int row = 2, column = getColumnNumber(columnForLoadingCode); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (codesInExcelFileFaculties.Contains(cellContent) == true)
                            {
                                duplicatesOfCodesInExcelFileFaculties.Add(row, cellContent);
                            }
                            else
                            {
                                codesInExcelFileFaculties.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                                missingValuesOfCodesInExcelFileFaculties.Add(row);
                            }
                        }
                        close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }


                    if (missingValuesOfNamesInExcelFileFaculties.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені назви факультетів в рядках: ");
                        foreach (int row in missingValuesOfNamesInExcelFileFaculties)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (missingValuesOfCodesInExcelFileFaculties.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені коди факультетів в рядках: ");
                        foreach (int row in missingValuesOfCodesInExcelFileFaculties)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (duplicatesOfNamesInExcelFileFaculties.Count != 0)
                    {
                        Console.WriteLine("Є дублікати назв факультетів:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfNamesInExcelFileFaculties)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                    }

                    if (duplicatesOfCodesInExcelFileFaculties.Count != 0)
                    {
                        Console.WriteLine("Є дублікати кодів факультетів:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfCodesInExcelFileFaculties)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                    }

                    //if (duplicatesOfNamesInExcelFileFaculties.Count == 0 && duplicatesOfCodesInExcelFileFaculties.Count == 0
                    //&& missingValuesOfCodesInExcelFileFaculties.Count == 0 && missingValuesOfNamesInExcelFileFaculties.Count == 0)
                    //{
                    try
                    {
                        ArrayList facultyNamesInDB = new ArrayList();
                        ArrayList facultyCodesInDB = new ArrayList();

                        MySqlConnection connection = DBUtils.GetDBConnection();
                        MySqlCommand mySqlCommand;
                        MySqlDataReader dataReader;

                        // Отримання даних з БД та порівняння з даними з Excel-файлу
                        // Якщо вони співпадають - немає сенсу для перезапису, інакше дані в БД перезаписуються

                        const string selectFaculties = "SELECT full_name, faculty_code FROM faculty";
                        connection.Open();
                        mySqlCommand = new MySqlCommand(selectFaculties, connection);
                        dataReader = mySqlCommand.ExecuteReader();

                        while (dataReader.Read())
                        {
                            //dataReader[0].ToString() - full_name
                            //dataReader[1].ToString() - faculty_code

                            facultyNamesInDB.Add(dataReader[0].ToString());
                            facultyCodesInDB.Add(dataReader[1].ToString());
                        }
                        connection.Close();


                        bool noSenseToReload = true;

                        foreach (string name in namesInExcelFileFaculties)
                        {
                            if (!facultyNamesInDB.Contains(name))
                            {
                                noSenseToReload = false;
                                break;
                            }
                        }

                        if (noSenseToReload)
                        {
                            foreach (string code in codesInExcelFileFaculties)
                            {
                                if (!facultyCodesInDB.Contains(code))
                                {
                                    noSenseToReload = false;
                                    break;
                                }
                            }
                        }

                        if (noSenseToReload == false)
                          {
                              Console.WriteLine("Є що змінювати");
                              try
                              {
                                  /* // очищення таблиці в БД
                                   connection.Open();
                                   const string truncateFaculties = "TRUNCATE TABLE faculty";
                                   mySqlCommand = new MySqlCommand(truncateFaculties, connection);
                                   mySqlCommand.ExecuteNonQuery();
                                   //Console.WriteLine("Очищено");
                                   connection.Close();
                                   */

                                  // перезапис таблиці в БД
                                  const string insertFaculties = "INSERT INTO faculty (full_name, faculty_code) VALUES (@FULL_NAME, @CODE)";

                                  connection.Open();
                                  for(int i = 0; i < namesInExcelFileFaculties.Count; i++)
                                  {
                                      mySqlCommand = new MySqlCommand(insertFaculties, connection);

                                      mySqlCommand.Parameters.AddWithValue("@FULL_NAME", namesInExcelFileFaculties[i]);
                                      mySqlCommand.Parameters.AddWithValue("@CODE", codesInExcelFileFaculties[i]);
                                      

                                      mySqlCommand.ExecuteNonQuery();
                                      //Console.WriteLine("code: " + codesInExcelFileFaculties[i] + "; name: " + namesInExcelFileFaculties[i]);
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
                    //}

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

