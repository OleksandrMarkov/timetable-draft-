using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class AuditoryTypes : ExcelFile
	{		
		
		
		ArrayList records = new ArrayList();
		ArrayList missingValues = new ArrayList();
		Dictionary <int, string> duplicates = new Dictionary<int, string>();
		int row = 1; // рядок, з якого починаються записи даних у файлі
		const char column = 'A'; // стовпець, з якого беруться дані
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		
		public AuditoryTypes(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
				
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				
				for(int col = getColumnNumber(column); row <= rowsCount; row++)
				{
					cellContent = getCellContent(row, col);
					//Console.WriteLine(cellContent);
					
					if(records.Contains(cellContent))
					{
						duplicates.Add(row, cellContent);
					}
					else
					{
						records.Add(cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValues.Add(row);
					}
				}				
				close();
			}
			catch(Exception ex)
			{
				reading = false;
				Console.WriteLine("Помилка при отриманні даних з файлу " + FileName + ". " + ex.Message);
			}			
		}
		
		public override void EvaluateData()
		{
			if (reading)
			{
				if (missingValues.Count != 0)
				{
					Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
	                foreach (int value in missingValues)
	                {
	                	Console.Write(value + "\t");
	                }	                
	                Console.WriteLine();
				}
				
				if(duplicates.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати типів аудиторій:");
					
					foreach (KeyValuePair<int, string> duplicate in duplicates)
	                {
	                	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
	                }
	                Console.WriteLine();
				}
			}		
			//Console.WriteLine("EvaluateData");
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{
					MySqlConnection connection  = DBUtils.GetDBConnection();
					MySqlCommand mySqlCommand;			
					const string insertTypes = "INSERT INTO auditory_type (auditory_type_name) VALUES (@TYPE)";
					
					connection.Open();
					foreach (string record in records)
					{
						mySqlCommand = new MySqlCommand(insertTypes, connection);
						mySqlCommand.Parameters.AddWithValue("@TYPE", record);
		                mySqlCommand.ExecuteNonQuery();
					}
					connection.Close();	
				}
				catch(Exception ex)
				{
					Console.WriteLine("Виникла помилка під час запису з файлу " + FileName + " до бази даних!");
				}	
			}			
			//Console.WriteLine("Load");
		}
	}
}
