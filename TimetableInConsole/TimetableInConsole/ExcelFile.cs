using System;
using System.Collections;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

using MySql.Data.MySqlClient;

namespace TimetableInConsole
{
	public class ExcelFile
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


        //protected char columnForLoading; // стовпець, що розглядається

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
                const string forbiddenSymbols = @"\" + "/:*?\"<>|";

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

        // завантаження до бази даних
        public void load()
        {
            switch (FileName)
            {
                case "TypesOfAuditories.xlsx":
                    ArrayList recordsInExcelFileAudTypes = new ArrayList();
                    ArrayList missingValuesInExcelFileAudTypes = new ArrayList();
                    Dictionary<int, string> duplicatesInExcelFileAudTypes = new Dictionary<int, string>();
                    char columnForLoading = 'A';

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
                        Console.WriteLine("В файлі " + FileName +  " є дублікати типів аудиторій:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesInExcelFileAudTypes)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }

                    if (missingValuesInExcelFileAudTypes.Count == 0 && duplicatesInExcelFileAudTypes.Count == 0)
                    {
                        try
                        {
                            ArrayList AudTypesInDB = new ArrayList();
                            //MySqlConnection connection = DBUtils.GetDBConnection();
                            MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
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
                        Console.WriteLine("В файлі " + FileName +  " є дублікати назв дисциплін:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesInExcelFileDisciplines)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }

                    //if (duplicatesInExcelFileDisciplines.Count == 0 && missingValuesInExcelFileDisciplines.Count == 0)
                    //{
                    try
                    {
                        ArrayList DisciplinesInDB = new ArrayList();
                        //MySqlConnection connection = DBUtils.GetDBConnection();
                        MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
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

                    const char columnForLoadingCode = 'D';
                    const char columnForLoadingName = 'B';

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
                        Console.WriteLine("В файлі " + FileName +  " є дублікати назв факультетів:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfNamesInExcelFileFaculties)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }

                    if (duplicatesOfCodesInExcelFileFaculties.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є дублікати кодів факультетів:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfCodesInExcelFileFaculties)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }

                    //if (duplicatesOfNamesInExcelFileFaculties.Count == 0 && duplicatesOfCodesInExcelFileFaculties.Count == 0
                    //&& missingValuesOfCodesInExcelFileFaculties.Count == 0 && missingValuesOfNamesInExcelFileFaculties.Count == 0)
                    //{
                    try
                    {
                        ArrayList facultyNamesInDB = new ArrayList();
                        ArrayList facultyCodesInDB = new ArrayList();

                        //MySqlConnection connection = DBUtils.GetDBConnection();
                        MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
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
                    ArrayList fullNamesInExcelFileDepartments = new ArrayList();
                    ArrayList shortNamesInExcelFileDepartments = new ArrayList();
                    ArrayList facultyCodesInExcelFileDepartments = new ArrayList();                    
                    
                    const char columnForLoadingFullName = 'A';
                    const char columnForLoadingShortName = 'B';
                    const char columnForLoadingFacultyCode = 'C';
                    
                    ArrayList missingValuesOfFullNamesInExcelFileDepartments = new ArrayList();
                    ArrayList missingValuesOfShortNamesInExcelFileDepartments = new ArrayList();
					ArrayList missingValuesOfFacultyCodesInExcelFileDepartments = new ArrayList();
                    
                    
                    Dictionary<int, string> duplicatesOfFullNamesInExcelFileDepartments = new Dictionary<int, string>();
                    Dictionary<int, string> duplicatesOfShortNamesInExcelFileDepartments = new Dictionary<int, string>();
                    
                    
                    try
                    {
                        open();
                        for (int row = 1, column = getColumnNumber(columnForLoadingFullName); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (fullNamesInExcelFileDepartments.Contains(cellContent) == true)
                            {
                               duplicatesOfFullNamesInExcelFileDepartments.Add(row, cellContent);
                            }
                            else
                            {
                                fullNamesInExcelFileDepartments.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                               missingValuesOfFullNamesInExcelFileDepartments.Add(row);
                            }
                        }

                        //Console.WriteLine();
                        
                        for (int row = 1, column = getColumnNumber(columnForLoadingShortName); row <= rowsCount; row++)
                        {
                            //Console.WriteLine(row);
                            //cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();                            
                            //Console.WriteLine("cellContent: " + cellContent);

                            // перевірка наявності дублікатів
                            if (shortNamesInExcelFileDepartments.Contains(cellContent) == true)
                            {
                            	duplicatesOfShortNamesInExcelFileDepartments.Add(row, cellContent);
                            }
                            else
                            {
                                shortNamesInExcelFileDepartments.Add(cellContent);
                            }
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                               missingValuesOfShortNamesInExcelFileDepartments.Add(row);
                            }
                        }

                               
                        //Console.WriteLine();
                        
                        for (int row = 1, column = getColumnNumber(columnForLoadingFacultyCode); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();                         
							//Console.WriteLine("cellContent: " + cellContent);
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                               missingValuesOfFacultyCodesInExcelFileDepartments.Add(row);
                            }
                            else
                            {
                            	facultyCodesInExcelFileDepartments.Add(cellContent);
                            }
                        }                        

                        close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }
                    
                    
                    if (missingValuesOfFullNamesInExcelFileDepartments.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені повні назви кафедр в рядках: ");
                        foreach (int row in missingValuesOfFullNamesInExcelFileDepartments)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }
                    if (missingValuesOfShortNamesInExcelFileDepartments.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені скорочені назви кафедр в рядках: ");
                        foreach (int row in missingValuesOfShortNamesInExcelFileDepartments)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }
                    if (missingValuesOfFacultyCodesInExcelFileDepartments.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені коди факультетів в рядках: ");
                        foreach (int row in missingValuesOfFacultyCodesInExcelFileDepartments)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }                    
                    
                    if (duplicatesOfFullNamesInExcelFileDepartments.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є дублікати повних назв кафедр:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfFullNamesInExcelFileDepartments)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }
                    if (duplicatesOfShortNamesInExcelFileDepartments.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є дублікати скорочених назв кафедр:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfShortNamesInExcelFileDepartments)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }

                    //if (duplicatesOfFullNamesInExcelFileDepartments.Count == 0 && duplicatesOfShortNamesInExcelFileDepartments.Count == 0
                    //&& missingValuesOfFullNamesInExcelFileDepartments.Count == 0 && missingValuesOfShortNamesInExcelFileDepartments.Count == 0 
                    //&& missingValuesOfFacultyCodesInExcelFileDepartments.Count == 0)
                    //{
                    try
                    {
                    	ArrayList departmentsInDB = new ArrayList();
                        ArrayList facultyIDInDB = new ArrayList();
                        
                        //ArrayList facultyCodesInDB = new ArrayList();
                        
                        MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
                        MySqlCommand mySqlCommand;
                        MySqlDataReader dataReader;
                        
	                    // Отримання даних з БД та порівняння з даними з Excel-файлу
                        // Якщо вони співпадають - немає сенсу для перезапису, інакше дані в БД перезаписуються
                        const string selectDepartments = "SELECT full_name FROM department";
                                    
                        connection.Open();
                        mySqlCommand = new MySqlCommand (selectDepartments, connection);
                        dataReader = mySqlCommand.ExecuteReader();                        
                        while (dataReader.Read())
                        {
                        	departmentsInDB.Add(dataReader[0].ToString());
                        	//Console.WriteLine(dataReader[0].ToString());                     	
                        }
                        connection.Close();
                        
                        bool noSenseToReload = true;
                        foreach (string name in fullNamesInExcelFileDepartments)
                        {
                        	if (!departmentsInDB.Contains(name))
                        	{
                        		noSenseToReload = false;
                        		break;
                        	}
                        }
                        
                        if (noSenseToReload == false)
                        {
                        	//Console.WriteLine("Є що змінювати");
                        	
	                        const string selectFacultyIDs = "SELECT faculty_id FROM faculty WHERE faculty_code = @CODE";
	                        
	                        const string insertDepartments = "INSERT INTO department (faculty_id, full_name, short_name) " +
	                        	"VALUES (@FACULTY_ID, @FULL_NAME, @SHORT_NAME)";
	                        
	                        connection.Open();
	                        for(int i = 0; i < facultyCodesInExcelFileDepartments.Count; i++)
	                        {
	                        	mySqlCommand = new MySqlCommand(selectFacultyIDs, connection);
	                        	mySqlCommand.Parameters.AddWithValue("@CODE", facultyCodesInExcelFileDepartments[i]);
	                        	mySqlCommand.ExecuteNonQuery();
	                        	
	                        	int facultyID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                        	
	                        	mySqlCommand = new MySqlCommand (insertDepartments, connection);
	                        	mySqlCommand.Parameters.AddWithValue("@FACULTY_ID", facultyID);
	                        	mySqlCommand.Parameters.AddWithValue("@FULL_NAME", fullNamesInExcelFileDepartments[i]);
	                        	mySqlCommand.Parameters.AddWithValue("@SHORT_NAME", shortNamesInExcelFileDepartments[i]);
	                        	mySqlCommand.ExecuteNonQuery();
	                        	/*Console.WriteLine(facultyID + " " + fullNamesInExcelFileDepartments[i]
	                        	                  + " " + shortNamesInExcelFileDepartments[i]);*/
	                        }
	                        connection.Close();
                        }
                        /*else
                        {
                        	Console.WriteLine("Нічого змінювати");
                        }*/
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
                    }
                    //}
                    break;
                    
                case "Teachers.xlsx":
                    ArrayList fioInExcelFileTeachers = new ArrayList();
                    ArrayList sexInExcelFileTeachers = new ArrayList();
                    ArrayList postsInExcelFileTeachers = new ArrayList();
                    ArrayList statusesInExcelFileTeachers = new ArrayList();

                    ArrayList departmentsInExcelFileTeachers = new ArrayList();
                    
                    const char columnForLoadingDepartments = 'A';
                    const char columnForLoadingFIO = 'J';
                    const char columnForLoadingSex = 'K';
                    const char columnForLoadingPosts = 'L';
                    const char columnForLoadingStatuses = 'M';
                    
                    const int firstRowInExcelFileTeachers = 8;
                    
                    
                    
                    ArrayList missingValuesOfDepartmentsInExcelFileTeachers = new ArrayList();
                    ArrayList missingValuesOfFIOInExcelFileTeachers = new ArrayList();
                    //ArrayList missingValuesOfSexInExcelFileTeachers = new ArrayList();
                    ArrayList missingValuesOfPostsInExcelFileTeachers = new ArrayList();
                    
                    //ArrayList missingValuesOfStatusesInExcelFileTeachers = new ArrayList();
                    
                    Dictionary<int, string> duplicatesOfFIOInExcelFileTeachers = new Dictionary<int, string>();
					Dictionary<int, string> wrongValuesOfSexInExcelFileTeachers = new Dictionary<int, string>();
                    
									
                    try
                    {
                        open();
                        // назви кафедр
                        for (int row = firstRowInExcelFileTeachers, column = getColumnNumber(columnForLoadingDepartments); row <= rowsCount/* + firstRowInExcelFileTeachers - 1*/; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
             
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	//throw new Exception("В файлі " + FileName + " пропущено назву кафедри в рядку: " + row); 
                            	missingValuesOfDepartmentsInExcelFileTeachers.Add(row);
                            }
                            else
                            {
                            	departmentsInExcelFileTeachers.Add(cellContent);	
                            }                            	                            
                        }
                        
                        // імена викладачів
                        for (int row = firstRowInExcelFileTeachers, column = getColumnNumber(columnForLoadingFIO); row <= rowsCount/* + firstRowInExcelFileTeachers - 1*/; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();                        
                            
             				if (fioInExcelFileTeachers.Contains(cellContent) == true)
                            {
                            	duplicatesOfFIOInExcelFileTeachers.Add(row, cellContent);                            	
                            }
             				if (string.IsNullOrEmpty(cellContent))
                            {
                              missingValuesOfFIOInExcelFileTeachers.Add(row);                             
                            }
                            fioInExcelFileTeachers.Add(cellContent);
                           
                        }
                        
                        // стать
                        for (int row = firstRowInExcelFileTeachers, column = getColumnNumber(columnForLoadingSex); row <= rowsCount /*+ firstRowInExcelFileTeachers - 1*/; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
                            
                            sexInExcelFileTeachers.Add(cellContent);
                            
                            if ( cellContent != "м" && cellContent != "ж")
                            {
                            	wrongValuesOfSexInExcelFileTeachers.Add(row, cellContent);
                            }
                        }
                        
                        // посади
                        for (int row = firstRowInExcelFileTeachers, column = getColumnNumber(columnForLoadingPosts); row <= rowsCount /*+ firstRowInExcelFileTeachers - 1*/; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                             
                            postsInExcelFileTeachers.Add(cellContent);
                            
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                              missingValuesOfPostsInExcelFileTeachers.Add(row);  
                            }
                        }
                        
                        // статус
                        for (int row = firstRowInExcelFileTeachers, column = getColumnNumber(columnForLoadingStatuses); row <= rowsCount /*+ firstRowInExcelFileTeachers - 1*/; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            
                            statusesInExcelFileTeachers.Add(cellContent);
                        }						
                    	close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }
                    
                    if (missingValuesOfDepartmentsInExcelFileTeachers.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені назви кафедр в рядках: ");
                        foreach (int row in missingValuesOfDepartmentsInExcelFileTeachers)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (duplicatesOfFIOInExcelFileTeachers.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є дублікати імен викладачів:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfFIOInExcelFileTeachers)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }
                    
                    if (missingValuesOfFIOInExcelFileTeachers.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені імена викладачів в рядках: ");
                        foreach (int row in missingValuesOfDepartmentsInExcelFileTeachers)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    if (wrongValuesOfSexInExcelFileTeachers.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є некоректні значення статі викладачів:");
                        foreach (KeyValuePair<int, string> wrongValue in wrongValuesOfSexInExcelFileTeachers)
                        {
                            Console.WriteLine("В рядку номер " + wrongValue.Key + ": " + wrongValue.Value);
                        }
                        Console.WriteLine();
                    }

                    if (missingValuesOfPostsInExcelFileTeachers.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені посади викладачів в рядках: ");
                        foreach (int row in missingValuesOfPostsInExcelFileTeachers)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }

                    // частина викладачів - без статусів
                    /*if (missingValuesOfStatusesInExcelFileTeachers.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущені статуси викладачів в рядках: ");
                        foreach (int row in missingValuesOfStatusesInExcelFileTeachers)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }*/         

                    
                       //if (missingValuesOfDepartmentsInExcelFileTeachers.Count == 0 && duplicatesOfFIOInExcelFileTeachers.Count == 0
                    //&& missingValuesOfFIOInExcelFileTeachers.Count == 0 && wrongValuesOfSexInExcelFileTeachers.Count == 0 
                    //&& missingValuesOfPostsInExcelFileTeachers.Count == 0 && missingValuesOfStatusesInExcelFileTeachers.Count == 0)
                    //{
                    
                    	//Console.WriteLine(rowsCount); 1190                  	
                    	/*Console.WriteLine("fio " +  fioInExcelFileTeachers.Count);
                    	Console.WriteLine("sex " + sexInExcelFileTeachers.Count);
                    	Console.WriteLine("post " + postsInExcelFileTeachers.Count);
                    	Console.WriteLine("status " + statusesInExcelFileTeachers.Count);*/
                    	
                    	// вставляємо БД все, а потім видаляємо "зайвих" викладачів
                    	// унікальність викладача: ФІО + id кафедри
                    	try
                    	{
                    		MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
                    		
	                    	const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT";
	                    	const string insertTeachers = "INSERT INTO teacher (department_id, full_name, sex, post, status) VALUES(@ID, @NAME, @SEX, @POST, @STATUS)";
                        	MySqlCommand mySqlCommand;
                        	//MySqlDataReader dataReader;
	                    	
	                        connection.Open();
	                       
	                    	for (int i = 0; i < rowsCount - firstRowInExcelFileTeachers + 1; i++)
	                    	{
	                    		mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
	                    		mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", departmentsInExcelFileTeachers[i]);
	                    		mySqlCommand.ExecuteNonQuery();
	                    		
	                    		int departmentID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                    		
	                    		mySqlCommand = new MySqlCommand(insertTeachers, connection);
	                    		mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
	                    		mySqlCommand.Parameters.AddWithValue("@NAME", fioInExcelFileTeachers[i]);
	                    		mySqlCommand.Parameters.AddWithValue("@SEX", sexInExcelFileTeachers[i]);
	                    		mySqlCommand.Parameters.AddWithValue("@POST", postsInExcelFileTeachers[i]);
	                    		mySqlCommand.Parameters.AddWithValue("@STATUS", statusesInExcelFileTeachers[i]);
	                    		mySqlCommand.ExecuteNonQuery();
	                    		/*Console.WriteLine(departmentID + " " + fioInExcelFileTeachers[i] + " " + sexInExcelFileTeachers[i] +
	                    		  " " + postsInExcelFileTeachers[i] + " " + statusesInExcelFileTeachers[i] );*/
	                    	}
	                    	
	                    	const string createTemporaryTable = "CREATE TEMPORARY TABLE teacher2 AS (SELECT * FROM teacher GROUP BY department_id, full_name)";
	                    	mySqlCommand = new MySqlCommand(createTemporaryTable, connection);
	                    	mySqlCommand.ExecuteNonQuery();
	                    	
	                    	const string deleteTrash = "DELETE FROM teacher WHERE teacher.teacher_id NOT IN (SELECT teacher2.teacher_id FROM teacher2)";
	                    	mySqlCommand = new MySqlCommand(deleteTrash, connection);
	                    	mySqlCommand.ExecuteNonQuery();
	                    	
	                    	//Console.WriteLine("мусор удален");
	                    	
	                    	connection.Close();
                    	}
                    	catch (Exception ex)
	                    {
	                        Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
	                    }
                    //}
                    
                    break;
                case "Auditories.xls":
                    
                    ArrayList namesInExcelFileAuditories = new ArrayList();
                    ArrayList typesInExcelFileAuditories = new ArrayList();
                    ArrayList departmentsInExcelFileAuditories = new ArrayList();
                    
                    ArrayList notUsedInExcelFileAuditories = new ArrayList();
                    ArrayList placesInExcelFileAuditories = new ArrayList();
                    
                    const char columnForLoadingNames = 'E';
                    const char columnForLoadingTypes = 'H';
                    const char columnForLoadingDepartmentNames = 'I';
                    const char columnForLoadingNotUsed = 'J';
                    const char columnForLoadingPlaces = 'G';
                     
                    const int firstRowInExcelFileAuditories = 2;
                    
                    
                    
                    ArrayList missingValuesOfNamesInExcelFileAuditories = new ArrayList();
                    ArrayList missingValuesOfTypesInExcelFileAuditories = new ArrayList();
                    ArrayList missingValuesOfDepartmentsInExcelFileAuditories = new ArrayList();
                    
                    
                    Dictionary<int, string> duplicatesOfNamesInExcelFileAuditories = new Dictionary<int, string>();

                    try
                    {
                        open();
                        
                        
						// назви аудиторій
                        for (int row = firstRowInExcelFileAuditories, column = getColumnNumber(columnForLoadingNames); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
             				
                            if(namesInExcelFileAuditories.Contains(cellContent))
                            {
                            	duplicatesOfNamesInExcelFileAuditories.Add(row, cellContent);
                            }
                            
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	//throw new Exception("В файлі " + FileName + " пропущено назву кафедри в рядку: " + row); 
                            	missingValuesOfNamesInExcelFileAuditories.Add(row);
                            }                            
                            else
                            {
                            	namesInExcelFileAuditories.Add(cellContent);	
                            }
                            
                        }
                        
                        // типи аудиторій
                        for (int row = firstRowInExcelFileAuditories, column = getColumnNumber(columnForLoadingTypes); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
             
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	missingValuesOfTypesInExcelFileAuditories.Add(row);
                            }
                            else
                            {
                            	typesInExcelFileAuditories.Add(cellContent);	
                            }                            	                            
                        }
                        
                        // кафедри, до яких належать аудиторії
                        for (int row = firstRowInExcelFileAuditories, column = getColumnNumber(columnForLoadingDepartmentNames); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
             
                            // перевірка наявності порожніх чарунок
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	missingValuesOfTypesInExcelFileAuditories.Add(row);
                            }
                            else
                            {
                            	departmentsInExcelFileAuditories.Add(cellContent);	
                            }                            	                            
                        }
                        
                        
                        // використання аудиторій
                        for (int row = firstRowInExcelFileAuditories, column = getColumnNumber(columnForLoadingNotUsed); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
 
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	notUsedInExcelFileAuditories.Add(false);
                            }
                            else
                            {
                            		notUsedInExcelFileAuditories.Add(true);
                            }                            	                            
                        }
                        
                        // скільки місць в аудиторії
                        for (int row = firstRowInExcelFileAuditories, column = getColumnNumber(columnForLoadingPlaces); row <= rowsCount; row++)
                        {
                            cellContent = ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
                            //Console.WriteLine("cellContent: " + cellContent);
             
                            if (string.IsNullOrEmpty(cellContent))
                            {
                            	placesInExcelFileAuditories.Add(0);
                            }
                            else
                            {
                            	placesInExcelFileAuditories.Add(Convert.ToInt32(cellContent));
                            }                            	                            
                        }
                        
                        close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                    }
					
                    /*Console.WriteLine(namesInExcelFileAuditories.Count);
                    Console.WriteLine(typesInExcelFileAuditories.Count);
                    Console.WriteLine(departmentsInExcelFileAuditories.Count);
                    Console.WriteLine(notUsedInExcelFileAuditories.Count);
                    Console.WriteLine(placesInExcelFileAuditories.Count);*/
                    
                   
                    if (duplicatesOfNamesInExcelFileAuditories.Count != 0)
                    {
                        Console.WriteLine("В файлі " + FileName +  " є дублікати назв кафедр:");
                        foreach (KeyValuePair<int, string> duplicate in duplicatesOfNamesInExcelFileAuditories)
                        {
                            Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                        }
                        Console.WriteLine();
                    }
                   
                    if (missingValuesOfNamesInExcelFileAuditories.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущено назви аудиторій в рядках: ");
                        foreach (int row in missingValuesOfNamesInExcelFileAuditories)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }
                    
                    if (missingValuesOfTypesInExcelFileAuditories.Count != 0)
                    {
                        Console.Write("В файлі " + FileName + " пропущено типи аудиторій в рядках: ");
                        foreach (int row in missingValuesOfTypesInExcelFileAuditories)
                        {
                            Console.Write(row + "\t");
                        }
                        Console.WriteLine();
                    }
                    /*if(duplicatesOfNamesInExcelFileAuditories.Count == 0 && missingValuesOfNamesInExcelFileAuditories.Count == 0 && missingValuesOfTypesInExcelFileAuditories.Count == 0)
                    {*/
						try
                    	{
                    		MySqlConnection connection = DBConnection.DBUtils.GetDBConnection();
                    		
	                    	const string selectAuditoryTypeID = "SELECT auditory_type_id FROM auditory_type WHERE auditory_type_name = @TYPE";
	                    	const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT_NAME";
	                    
	                    	const string insertAuditories = "INSERT INTO auditory (department_id, auditory_name, not_used, type_auditory, count_of_places) VALUES(@ID, @AUDITORY_NAME, @NOT_USED, @TYPE_ID, @COUNT)";
                        	MySqlCommand mySqlCommand;
	                    	
	                        connection.Open();
	                       
	                    	for (int i = 0; i < rowsCount - firstRowInExcelFileAuditories + 1; i++)
	                    	{
	                    		mySqlCommand = new MySqlCommand(selectAuditoryTypeID, connection);
	                    		mySqlCommand.Parameters.AddWithValue("@TYPE", typesInExcelFileAuditories[i]);
	                    		mySqlCommand.ExecuteNonQuery();
	                    		
	                    		int auditoryTypeID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                    		
	                    		mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
	                    		mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_NAME", departmentsInExcelFileAuditories[i]);
	                    		mySqlCommand.ExecuteNonQuery();
	                    		
	                    		int departmentID = Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                    		
	                    		mySqlCommand = new MySqlCommand(insertAuditories, connection);
	                    		
	                    		mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
	                    		mySqlCommand.Parameters.AddWithValue("@AUDITORY_NAME", namesInExcelFileAuditories[i]);
	                    		mySqlCommand.Parameters.AddWithValue("@NOT_USED", notUsedInExcelFileAuditories[i]);
	                    		mySqlCommand.Parameters.AddWithValue("@TYPE_ID", auditoryTypeID);
	                    		mySqlCommand.Parameters.AddWithValue("@COUNT", placesInExcelFileAuditories[i]);
	                    		mySqlCommand.ExecuteNonQuery();
	                    		
	                    		Console.WriteLine(departmentID + " " + namesInExcelFileAuditories[i] + " " + notUsedInExcelFileAuditories[i] +
	                    		  " " + auditoryTypeID + " " + placesInExcelFileAuditories[i] );
	                    	}
	                    		                    	
	                    	connection.Close();
                    	}
                    	catch (Exception ex)
	                    {
	                        Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
	                    }
            //}
                                       
                    break;
                case "StudyGroups.xlsx":
                    break;
                default:
                    throw new Exception("Невідомий файл");
            }
        }
	}
}
