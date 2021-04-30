using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class StudyGroups : ExcelFile
	{		
		public StudyGroups(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
		
		ArrayList departments = new ArrayList();
		ArrayList names = new ArrayList();
		ArrayList courseNumbers = new ArrayList();
		ArrayList countOfStudents = new ArrayList();		
		
		ArrayList codes = new ArrayList();
		
		ArrayList missingValuesOfDepartments = new ArrayList();
		ArrayList missingValuesOfNames = new ArrayList();		
		
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		
		const char departmentsColumn = 'A'; // стовпець, з якого беруться назви кафедр (скорочені)
		const char namesColumn = 'B'; // стовпець, з якого беруться назви груп
		const char countOfStudentsColumn = 'C'; // стовпець, з якого беруться кількості студентів в групах
		const char courseNumbersColumn = 'D'; // стовпець, з якого беруться номери курсів
		
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
		
		int row = 1; // рядок, з якого починаються записи даних у файлі
	
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				// назви кафедр (скорочені)
				for(int col = getColumnNumber(departmentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
											
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfDepartments.Add(i);
					}					
					else
					{
						cellContent = cellContent.Replace("*", "");
						departments.Add(cellContent);
					}
				}

				// назви груп
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
											
					if(string.IsNullOrEmpty(cellContent))
					{
						missingValuesOfNames.Add(i);
					}
					else
					{
						if(names.Contains(cellContent))
						{
							duplicatesOfNames.Add(i, cellContent);
						}
						
						names.Add(cellContent);
						//коди груп
						int hyphen = cellContent.IndexOf('-');
						codes.Add(cellContent.Substring(0, hyphen + 2));
					}
				}
				
				// курси
				for(int col = getColumnNumber(courseNumbersColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					courseNumbers.Add(cellContent);
				}
				
				// кількості студентів в групах
				for(int col = getColumnNumber(countOfStudentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
											
					if(string.IsNullOrEmpty(cellContent))
					{
						countOfStudents.Add(null);
					}					
					else
					{
						countOfStudents.Add(Convert.ToInt32(cellContent));
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
				/*Console.WriteLine(names.Count);
				Console.WriteLine(codes.Count);
				Console.WriteLine(courseNumbers.Count);
				Console.WriteLine(countOfStudents.Count);
				*/
				
				/*if(missingValuesOfNames.Count != 0)
				{
						Console.Write("В файлі " + FileName + " пропущено назви груп в рядках: ");
		                foreach (int value in missingValuesOfNames)
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

                if (duplicatesOfNames.Count != 0)
                {
                	Console.WriteLine("В файлі " + FileName +  " є дублікати назв груп:");
                    foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
                    {
                    	Console.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value);
                    }
                        Console.WriteLine();
                }*/					
			}
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
                {
	                MySqlConnection connection = DBUtils.GetDBConnection();
	                    		
		            //const string selectDepartmentID = "SELECT department_id FROM department WHERE short_name = @DEPARTMENT";
		            const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT";
		            const string insertStudyGroups = "INSERT INTO study_group (department_id, study_group_code, full_name, course_number, count_of_students) VALUES(@ID, @CODE, @NAME, @COURSE, @COUNT)";
	                MySqlCommand mySqlCommand;
		                    	
		            connection.Open();
		                       
		            for (int i = 0; i < rowsCount; i++)
		            {
			            mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
			            mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", departments[i]);
			            mySqlCommand.ExecuteNonQuery();
			                    		
			            int departmentID =  Convert.ToInt32( mySqlCommand.ExecuteScalar().ToString() );
			           // Console.WriteLine(departmentID);
			            
						mySqlCommand = new MySqlCommand(insertStudyGroups, connection);
						mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
						mySqlCommand.Parameters.AddWithValue("@CODE", codes[i]);
						mySqlCommand.Parameters.AddWithValue("@NAME", names[i]);
						mySqlCommand.Parameters.AddWithValue("@COURSE", courseNumbers[i]);
						mySqlCommand.Parameters.AddWithValue("@COUNT", countOfStudents[i]);	                    		
						mySqlCommand.ExecuteNonQuery();
					
						/*Console.WriteLine(departmentID + " " + codes[i] + " " + names[i] +
						" " + countOfStudents[i]);*/
	            	 }
	                 connection.Close();
	                 
	                 Console.WriteLine("Study groups are loaded");
                 }
				catch (Exception ex)
				{
					Console.WriteLine("Помилка при завантаженні даних з файлу " + FileName + "\n" + ex.Message);
				}				
			}
		}
	}
}
