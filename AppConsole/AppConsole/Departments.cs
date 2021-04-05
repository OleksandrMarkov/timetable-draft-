using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class Departments : ExcelFile
	{
		ArrayList full_names = new ArrayList();
		ArrayList short_names = new ArrayList();
		ArrayList faculty_codes = new ArrayList();
		
		ArrayList missingValuesOfFullNames = new ArrayList();
		ArrayList missingValuesOfShortNames = new ArrayList();
		ArrayList missingValuesOfFacultyCodes = new ArrayList();
		
		
		Dictionary <int, string> duplicatesOfFullNames = new Dictionary<int, string>();
		Dictionary <int, string> duplicatesOfShortNames = new Dictionary<int, string>();
		int row = 1; // рядок, з якого починаються записи даних у файлі
		
		const char fullNamesColumn = 'A'; // стовпець, з якого беруться назви кафедр
		const char shortNamesColumn = 'B'; // стовпець, з якого беруться скорочені назви кафедр
		const char facultyCodesColumn = 'C'; // стовпець, з якого беруться коди факультетів
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		
		public Departments(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				for(int col = getColumnNumber(fullNamesColumn); row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
					
					if(full_names.Contains(cellContent))
					{
						duplicatesOfFullNames.Add(row, cellContent);
					}
					else
					{
						full_names.Add(cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfFullNames.Add(row);
					}
				}
				
				for(int col = getColumnNumber(shortNamesColumn), row = 1; row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
					
					if(short_names.Contains(cellContent))
					{
						duplicatesOfShortNames.Add(row, cellContent);
					}
					else
					{
						short_names.Add(cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfShortNames.Add(row);
					}				
				}
				
				for(int col = getColumnNumber(facultyCodesColumn), row = 1; row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
										
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfFacultyCodes.Add(row);
					}
					else
					{
						faculty_codes.Add(cellContent);
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
				if(missingValuesOfFullNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfFullNames)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(missingValuesOfShortNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfShortNames)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(missingValuesOfFacultyCodes.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfFacultyCodes)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(duplicatesOfFullNames.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати назв кафедр:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfFullNames)
	                {
	                	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
	                }	
					Console.WriteLine();
				}
				
				if(duplicatesOfShortNames.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати скорочених назв кафедр:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfShortNames)
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
					
					const string selectFacultyIDs = "SELECT faculty_id FROM faculty WHERE faculty_code = @CODE";
									
					const string insertDepartments = "INSERT INTO department (faculty_id, full_name, short_name) "
					+ "VALUES (@FACULTY_ID, @FULL_NAME, @SHORT_NAME)";
					
					connection.Open();
					
					for(int i = 0; i < faculty_codes.Count; i++)
					{
						mySqlCommand = new MySqlCommand(selectFacultyIDs, connection);
						mySqlCommand.Parameters.AddWithValue("@CODE", faculty_codes[i]);
						
						mySqlCommand.ExecuteNonQuery();
						
						int facultyID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
						
						mySqlCommand = new MySqlCommand(insertDepartments, connection);
						
						mySqlCommand.Parameters.AddWithValue("@FACULTY_ID", facultyID);
	                    mySqlCommand.Parameters.AddWithValue("@FULL_NAME", full_names[i]);
	                    mySqlCommand.Parameters.AddWithValue("@SHORT_NAME", short_names[i]);
	                    mySqlCommand.ExecuteNonQuery();
					}
					
					connection.Close();
				}
				catch(Exception ex)
				{
					Console.WriteLine("Виникла помилка під час запису з файлу " + FileName + " до бази даних!");
				}
			}
		}
		
	}
}
