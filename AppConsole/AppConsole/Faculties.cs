using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class Faculties : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList codes = new ArrayList();
		
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfCodes = new ArrayList();
		
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		Dictionary <int, string> duplicatesOfCodes = new Dictionary<int, string>();
		int row = 2; // рядок, з якого починаються записи даних у файлі
		
		const char namesColumn = 'B'; // стовпець, з якого беруться назви факультетів
		const char codesColumn = 'D'; // стовпець, з якого беруться коди факультетів
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		
		public Faculties(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
			
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				for(int col = getColumnNumber(namesColumn); row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
					
					if(names.Contains(cellContent))
					{
						duplicatesOfNames.Add(row, cellContent);
					}
					
					else
					{
						names.Add(cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfNames.Add(row);
					}
				}
				
				for(int col = getColumnNumber(codesColumn), row = 2; row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
					
					if(codes.Contains(cellContent))
					{
						duplicatesOfCodes.Add(row, cellContent);
					}
					else
					{
						codes.Add(cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfCodes.Add(row);
					}				
				}
				
				close();
			}
			catch (Exception ex)
            {
				reading = false;
            	Console.WriteLine("Помилка при отриманні даних з файлу " + FileName + " " + ex.Message);
            }
			
		}
		
		public override void EvaluateData()
		{
			if(reading)
			{
				if(missingValuesOfNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfNames)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(duplicatesOfNames.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати назв факультетів:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
	                {
	                	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
	                }
	
					Console.WriteLine();
				}	
	
				if(missingValuesOfCodes.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfCodes)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(duplicatesOfCodes.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати кодів факультетів:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfCodes)
	                {
	                	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
	                }
	
					Console.WriteLine();
				}	
			}						
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{
					MySqlConnection connection = DBUtils.GetDBConnection();
					MySqlCommand mySqlCommand;			
					const string insertFaculties = "INSERT INTO faculty (full_name, faculty_code) VALUES (@FULL_NAME, @CODE)";
					
					connection.Open();
					
					for(int i = 0; i < names.Count; i++)
					{
						mySqlCommand = new MySqlCommand(insertFaculties, connection);
						
						mySqlCommand.Parameters.AddWithValue("@FULL_NAME", names[i]);
                        mySqlCommand.Parameters.AddWithValue("@CODE", codes[i]);
                        
                        mySqlCommand.ExecuteNonQuery();
                        
                        //Console.WriteLine("code: " + codes[i] + "; name: " + names[i]);
					}
					
					connection.Close();
					
				}
				catch(Exception ex)
				{
					Console.WriteLine("Виникла помилка під час запису з файлу " + FileName + " до бази даних!" + "\n" + ex.Message);
				}
			
			}			
		}		
	}
}
