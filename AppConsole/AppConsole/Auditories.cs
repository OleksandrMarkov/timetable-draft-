using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{

	public class Auditories : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList types = new ArrayList();
		ArrayList departments = new ArrayList();
		ArrayList notUsed = new ArrayList();
		ArrayList places = new ArrayList();
		
		
		const char namesColumn = 'E'; // стовпець, з якого беруться назви аудиторій
		const char typesColumn = 'H'; // стовпець, з якого беруться типи аудиторій
		const char departmentsColumn = 'I'; // стовпець, з якого беруться кафедри
		
		const char notUsedColumn = 'J'; // стовпець, з якого беруться значення, чи використовуються аудиторії 
		const char placesColumn = 'G'; // стовпець, з якого беруться кількості місць в аудиторіях
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		int row = 2; // рядок, з якого починаються записи даних у файлі
		
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfTypes = new ArrayList();
		ArrayList missingValuesOfDepartments = new ArrayList();
		
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		
		public Auditories(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
		
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				// назви аудиторій
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
						
					if (names.Contains(cellContent))
					{
						duplicatesOfNames.Add(i, cellContent);
					}
						
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfNames.Add(i);
					}
						
					else
					{
						names.Add(cellContent);
					}
				}

				// типи аудиторій
				for(int col = getColumnNumber(typesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
												
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfTypes.Add(i);
					}						
					else
					{
						types.Add(cellContent);
					}
				}
				
				// кафедри, яким належать аудиторії
				for(int col = getColumnNumber(departmentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
												
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfDepartments.Add(i);
					}						
					else
					{
						departments.Add(cellContent);
					}
				}
				

				// аудиторії, які використовуються
				for(int col = getColumnNumber(notUsedColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
												
					if(string.IsNullOrEmpty(cellContent))
					{
						notUsed.Add(false);
					}						
					else
					{
						notUsed.Add(true);
					}
				}

				// скільки місць в аудиторіях
				for(int col = getColumnNumber(placesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
												
					if(string.IsNullOrEmpty(cellContent))
					{
						places.Add(0);
					}						
					else
					{
						places.Add(Convert.ToInt32(cellContent));
					}
				}				
				
				close();
				
			}
			catch (Exception ex)
		    {
				reading = false;
		    	Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
		    }	
		}

		public override void EvaluateData()
		{
			if(reading)
			{
				if(missingValuesOfNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " пропущено назви аудиторій в рядках: ");
		                foreach (int value in missingValuesOfNames)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}

				if(missingValuesOfTypes.Count != 0)
				{
					Console.Write("В файлі " + FileName + " пропущено типи аудиторій в рядках: ");
		            foreach (int value in missingValuesOfTypes)
		            {
		            	Console.Write(value + "\t");
		            }	                
		            Console.WriteLine();
				}				

				if(missingValuesOfDepartments.Count != 0)
				{
					Console.Write("В файлі " + FileName + " пропущено назви кафедр в рядках: ");
		            foreach (int value in missingValuesOfDepartments)
		            {
		            	Console.Write(value + "\t");
		            }	                
		            Console.WriteLine();
				}
				
				if(duplicatesOfNames.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати назв аудиторій:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
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
					
					const string selectAuditoryTypeID = "SELECT auditory_type_id FROM auditory_type WHERE auditory_type_name = @TYPE";
					const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT_NAME";					
					
					const string insertAuditories = "INSERT INTO auditory (department_id, auditory_name, not_used, type_auditory, count_of_places) VALUES(@ID, @AUDITORY_NAME, @NOT_USED, @TYPE_ID, @COUNT)";
					
					connection.Open();
					
					for(int i = 0; i < rowsCount - row + 1; i++)
					{
						mySqlCommand = new MySqlCommand(selectAuditoryTypeID, connection);
	                    mySqlCommand.Parameters.AddWithValue("@TYPE", types[i]);
	                    mySqlCommand.ExecuteNonQuery();
	                    		
	                    int auditoryTypeID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                    		
	                    mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
	                    mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_NAME", departments[i]);
	                    mySqlCommand.ExecuteNonQuery();
	                    		
	                    int departmentID = Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
	                    		
	                    mySqlCommand = new MySqlCommand(insertAuditories, connection);
	                    		
	                    mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
	                    mySqlCommand.Parameters.AddWithValue("@AUDITORY_NAME", names[i]);
	                    mySqlCommand.Parameters.AddWithValue("@NOT_USED", notUsed[i]);
	                    mySqlCommand.Parameters.AddWithValue("@TYPE_ID", auditoryTypeID);
	                    mySqlCommand.Parameters.AddWithValue("@COUNT", places[i]);
	                    mySqlCommand.ExecuteNonQuery();	
					}
					
					connection.Close();
					
					Console.WriteLine("!!!!");
				}
				
                catch (Exception ex)
	            {
	            	Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
	            }								
			}
		}
		
	}
}
