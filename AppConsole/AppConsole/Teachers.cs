using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{

	public class Teachers : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList sex = new ArrayList();
		ArrayList posts = new ArrayList();
		ArrayList statuses = new ArrayList();
		
		ArrayList departments = new ArrayList();
		
		const char departmentsColumn = 'A'; // стовпець, з якого беруться назви кафедр
		const char namesColumn = 'J'; // стовпець, з якого беруться ПІБ викладачів
		const char sexColumn = 'K'; // стовпець, з якого беруться статі викладачів (м/ж)
		const char postsColumn = 'L'; // стовпець, з якого беруться посади викладачів
		const char statusesColumn = 'M'; // стовпець, з якого беруться звання (статуси) викладачів
		
		int row = 8; // рядок, з якого починаються записи даних у файлі
		
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfPosts = new ArrayList();
		ArrayList missingValuesOfDepartments = new ArrayList();
		
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		Dictionary <int, string> wrongValuesOfSex = new Dictionary<int, string>();
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		public Teachers(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
		
		/* на відміну від інших файлів, в БД вставлються УСІ записи,
		а вже потім звідти видаляється зайве*/
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				
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
				
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					if(names.Contains(cellContent))
					{
						duplicatesOfNames.Add(i, cellContent);
					}
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfNames.Add(i);
					}
					
					names.Add(cellContent);
				}
				
				for(int col = getColumnNumber(sexColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					sex.Add(cellContent);
					
					if ( cellContent != "м" && cellContent != "ж")
                    {
						wrongValuesOfSex.Add(i, cellContent);
					}
				}
				
				for(int col = getColumnNumber(postsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					posts.Add(cellContent);
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfPosts.Add(i);
					}
				}
				
				for(int col = getColumnNumber(postsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					statuses.Add(cellContent);
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
				if (missingValuesOfDepartments.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfDepartments)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if (missingValuesOfNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfNames)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if (missingValuesOfPosts.Count != 0)
				{
						Console.Write("В файлі " + FileName + " є пропуски в рядках: ");
		                foreach (int value in missingValuesOfPosts)
		                {
		                	Console.Write(value + "\t");
		                }	                
		                Console.WriteLine();
				}
				
				if(duplicatesOfNames.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є дублікати імен викладачів:");
		
					foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
	                {
	                	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
	                }	
					Console.WriteLine();
				}
				
				if(wrongValuesOfSex.Count != 0)
				{
					Console.WriteLine("В файлі " + FileName +  " є некоректні значення статі викладачів:");
		
					foreach (KeyValuePair<int, string> wrongValue in wrongValuesOfSex)
	                {
	                	Console.WriteLine("В рядку номер " + wrongValue.Key + ": " + wrongValue.Value);
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
					
					const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT";
					const string insertTeachers = "INSERT INTO teacher (department_id, full_name, sex, post, status) " +
						"VALUES(@ID, @NAME, @SEX, @POST, @STATUS)";
					
					connection.Open();
					
					//Console.WriteLine(row);
					
					for(int i = 0; i < rowsCount - row + 1; i++)
					{
						mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
						mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", departments[i]);
						mySqlCommand.ExecuteNonQuery();
						
						int departmentID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString());
						
						mySqlCommand = new MySqlCommand(insertTeachers, connection);
						
						mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
						mySqlCommand.Parameters.AddWithValue("@NAME", names[i]);
						mySqlCommand.Parameters.AddWithValue("@SEX", sex[i]);
						mySqlCommand.Parameters.AddWithValue("@POST", posts[i]);
						mySqlCommand.Parameters.AddWithValue("@STATUS", statuses[i]);
						mySqlCommand.ExecuteNonQuery();
						
					}
					
					const string createTemporaryTable = "CREATE TEMPORARY TABLE teacher2 AS (SELECT * FROM teacher GROUP BY department_id, full_name)";
					mySqlCommand = new MySqlCommand(createTemporaryTable, connection);
	                mySqlCommand.ExecuteNonQuery();
					
	                const string deleteTrash = "DELETE FROM teacher WHERE teacher.teacher_id NOT IN (SELECT teacher2.teacher_id FROM teacher2)";
	                mySqlCommand = new MySqlCommand(deleteTrash, connection);
	                mySqlCommand.ExecuteNonQuery();
					connection.Close();
					
					//Console.WriteLine("IT IS LOADED!");
				}
				catch (Exception ex)
	            {
	            	Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
	            }
			}
		}		
	}
}