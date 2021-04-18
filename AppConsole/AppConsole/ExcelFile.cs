using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace AppConsole
{
	public class ExcelFile
	{
		
		protected string fileName;
		protected const string directoryName = @"E:\BACHELORS WORK\TIMETABLE\data";
		
		protected Excel.Application app;
		protected Excel.Workbook workbook;
		protected Excel.Worksheet worksheet;
		protected Excel.Range range;
		
		protected int rowsCount; // к-ть рядків з даними
		protected string cellContent;
		
		protected string wrongFileName; // при спробі відкрити файл, якого немає в директорії, сюди запишеться його ім'я
		
		public string FileName
        {
            get
            {
                return fileName;
            }

            set
            {
                fileName = null; // якщо ім'я файлу буде некоректним, в змінній "fileName" залишиться null


                // Перевірка наявності заборонених символів в імені Excel-файлу
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
                // + Перевірка розширення файлу (".xlsx" або ".xls")
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
                return String.Concat(directoryName, @"\", fileName);
            }
        }
		
        public ExcelFile(){}
        
        public ExcelFile(string fileName)
        {
            FileName = fileName;
        }
        
        //відкриття файлу
        public void open(int sheetNumber)
        {
            if (fileName == null)
            {
                Console.WriteLine("Помилка відкриття файлу \"" + wrongFileName +
                    "\". Перевірте правильність імені файлу та повторіть спробу.");
            }
            else
            {
                app = new Excel.Application();

                const int readingMode = 0; // режим читання файлу
				//const int sheetNumber = 1; // інформація в файлах знаходиться на першому листі
                try
                {
                    workbook = app.Workbooks.Open(FullPathToFile, readingMode, true); // відкриття книги 
                    worksheet = (Excel.Worksheet)workbook.Worksheets[sheetNumber]; // відкриття листа даних                

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
            if(!File.Exists(FullPathToFile))
            {
                Console.WriteLine("Файл не знайдено: " + "\"" + wrongFileName + "\"");
                return false;
            }
            return true;
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

        // отримання номера стовпця, що зчитується: 'A' - 1, 'B' - 2, 'C' - 3, 'D' - 4, ...		
        protected int getColumnNumber(char column)
        {
            if ((int)column >= 'A' && (int)column <= 'Z')
            {
                return (int)column - 64;
            }
            else
            {
                throw new Exception("Некоректне ім'я стовпця: " + column + ". Неможливо зчитати дані з файлу " + FileName);
            }
        }
        
        protected string getCellContent(int row, int column)
        {
        	return ((Excel.Range)worksheet.Cells[row, column]).Text.ToString();
        }        
        
        // Паттерн "ШАБЛОННИЙ МЕТОД"
        public void SendDataToDB()
		{
			ReadFromExcelFile();
			EvaluateData();
			Load();
		}
		
        public virtual void ReadFromExcelFile(){}
        public virtual void EvaluateData(){}
        public virtual void Load(){}
	}
}
