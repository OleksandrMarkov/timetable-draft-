using System;
using System.IO;
using System.Diagnostics;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataCollectionApp
{
	public class ExcelFile
	{
		protected string fileName;
		protected const string directoryName = @"E:\BACHELORS WORK\TIMETABLE\data";
		protected Excel.Application app;
		protected Excel.Workbook workbook;
		protected Excel.Worksheet worksheet;
		protected Excel.Range range;
		protected int rowsCount;
		protected string cellContent;	
		protected string wrongFileName;		
		public string FileName
        {
            get { return fileName; }

            set
            {
                fileName = null; 
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

                if (containsForbiddenSymbols == false && (value.EndsWith(".xlsx") || value.EndsWith(".xls")))
                {
                    fileName = value;
                }
                else
                {
                    wrongFileName = value;
                    MessageBox.Show("Некоректне ім'я або розширення файлу \"" + wrongFileName +
                        "\". Ім'я файлу не повинно містити наступних символів: " + forbiddenSymbols +
                        "\nРозширення файлу повинно бути \".xls\" або \".xlsx\"\n");
                }
            }
        }	
        public string FullPathToFile
        {
            get { return String.Concat(directoryName, @"\", fileName); }
        }	
        public ExcelFile(){}
        public ExcelFile(string fileName)
        { FileName = fileName; }
        public void open(int sheetNumber)
        {
            if (fileName == null)
            {
            	MessageBox.Show("Помилка відкриття файлу \"" + wrongFileName +
                    "\". Перевірте правильність імені файлу та повторіть спробу.");
            }
            else
            {
                app = new Excel.Application();

                const int readingMode = 0;
                try
                {
                    workbook = app.Workbooks.Open(FullPathToFile, readingMode, true);
                    worksheet = (Excel.Worksheet)workbook.Worksheets[sheetNumber];               

                    range = worksheet.UsedRange;
                    rowsCount = range.Rows.Count;

                }
                catch (Exception ex)
                {
                	MessageBox.Show("Помилка відкриття файлу \"" + wrongFileName +
                        "\". Перевірте правильність імені файлу та повторіть спробу.");
                }
            }
        }
        public void openForViewing()
        {	Process.Start(FullPathToFile); }   
        public string LastWriteTime
        {
        	get
        	{
        		FileInfo f = new FileInfo(FullPathToFile);
        		f.Refresh();	
        		return f.LastWriteTime.ToString();       		
        	}
        }       
        public bool exists()
        {
            if(!File.Exists(FullPathToFile))
            {
            	MessageBox.Show("Файл не знайдено: " + "\"" + wrongFileName + "\"");
                return false;
            }
            return true;
        }
        public void close()
        {
            try
            {
                workbook.Close(false);
                app.Quit();
            }
            catch (Exception ex)
            {
            	MessageBox.Show("Помилка при закритті файлу: \"" + wrongFileName + "\".");
            }
        }   
        protected int getColumnNumber(char column)
        {
            if ((int)column >= 'A' && (int)column <= 'Z')
            { return (int)column - 64; }
            else
            {
            	MessageBox.Show("Некоректне ім'я стовпця: " + column + ". Неможливо зчитати дані з файлу " + FileName);
            	return -1;
            }
        }        
        
		protected string getCellContent(int row, int column)
        { return ((Excel.Range)worksheet.Cells[row, column]).Text.ToString(); }        
        
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