using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

using System.IO;

using System.Windows; // for messageBoxes

namespace DataCollectionApp
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
		
		/* на відміну від завантаження інших файлів, в БД вставлються УСІ записи,
		а вже потім звідти видаляється зайве*/
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				
				// назви кафедр
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
				
				// ПІБ викладачів
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
				
				// стать викладачів
				for(int col = getColumnNumber(sexColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					sex.Add(cellContent);
					
					if ( cellContent != "м" && cellContent != "ж")
                    {
						wrongValuesOfSex.Add(i, cellContent);
					}
				}
				
				//посади викладачів
				for(int col = getColumnNumber(postsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					posts.Add(cellContent);
					
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfPosts.Add(i);
					}
				}
				
				// звання (статуси) викладачів
				for(int col = getColumnNumber(statusesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					statuses.Add(cellContent);
				}
				
				close();
			}
			catch (Exception ex)
            {
				reading = false;
            	MessageBox.Show("Помилка при отриманні даних з файлу " + FileName + " " + ex.Message);
            }
		}
		
		public override void EvaluateData()
		{
			if(reading)
			{
				const string path = @"E:\BACHELORS WORK\TIMETABLE\DataCollectionApp\BugsReport.txt";
				if (missingValuesOfDepartments.Count != 0 || missingValuesOfNames.Count != 0 ||
				   missingValuesOfPosts.Count !=0 || duplicatesOfNames.Count != 0 ||
				  wrongValuesOfSex.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("ВИКЛАДАЧІ.");
						sw.WriteLine("Файл: " + FileName);
					}
				}
				
				if (missingValuesOfDepartments.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви кафедр в рядках: ");
						foreach (int value in missingValuesOfDepartments)
						{
							sw.Write(value + "|");
						}
						sw.WriteLine();
					}
				}
				
				if (missingValuesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено ПІБ викладачів в рядках: ");
						foreach (int value in missingValuesOfNames)
						{
							sw.Write(value + "|");
						}
						sw.WriteLine();
					}
				}
				
				if (missingValuesOfPosts.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено посади викладачів в рядках: ");
						foreach (int value in missingValuesOfPosts)
						{
							sw.Write(value + "|");
						}
						sw.WriteLine();
					}
				}
				
				if(duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати ПІБ викладачів: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
						{
							sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
						}
						sw.WriteLine();
					}
				}
				
				if(wrongValuesOfSex.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є некоректні значення статі викладачів: ");
						foreach (KeyValuePair<int, string> wrongValue in wrongValuesOfSex)
						{
							sw.WriteLine("В рядку номер " + wrongValue.Key + ": " + wrongValue.Value);
						}
						sw.WriteLine();
					}
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
	                const string insertTeachers = "INSERT INTO teacher (department_id, full_name, sex, post, status) VALUES(@ID, @NAME, @SEX, @POST, @STATUS)";
					
					
					connection.Open();
					
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
					
				}
				catch (Exception ex)
	            {
					MessageBox.Show("Виникла помилка під час завантаження даних про викладачів з файлу " + FileName + " до бази даних!" + "\n " + ex.Message);
				}
			}
		}		
	}
}