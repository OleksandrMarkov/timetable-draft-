using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Windows;
namespace DataCollectionApp
{
	public class Auditories : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList types = new ArrayList();
		ArrayList departments = new ArrayList();
		ArrayList notUsed = new ArrayList();
		ArrayList places = new ArrayList();
		ArrayList corpsNumbers = new ArrayList();
		
		const char namesColumn = 'E';
		const char typesColumn = 'H';
		const char departmentsColumn = 'I';
		const char notUsedColumn = 'J';
		const char placesColumn = 'G';
		bool reading = true;	
		int row = 2;
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfTypes = new ArrayList();
		ArrayList missingValuesOfDepartments = new ArrayList();
		
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		
		public Auditories(string fileName): base(fileName)
		{ this.fileName = fileName; }
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					string trimmedCellContent = cellContent.TrimStart('0');
					if (!string.IsNullOrEmpty(cellContent) && string.IsNullOrEmpty(trimmedCellContent))
					{ trimmedCellContent = "0"; }						
					if (names.Contains(trimmedCellContent))
					{ duplicatesOfNames.Add(i, trimmedCellContent); }
					if(string.IsNullOrEmpty(trimmedCellContent))
					{ missingValuesOfNames.Add(i); }
						
					else
					{	
						names.Add(trimmedCellContent);
						if (trimmedCellContent.StartsWith("4") && trimmedCellContent.Length > 2 &&
						    Char.IsDigit(trimmedCellContent[1]) && Char.IsDigit(trimmedCellContent[2]))
						{ corpsNumbers.Add(4);}
						else
						{
							if (trimmedCellContent.StartsWith("5") && trimmedCellContent.Length > 2 &&
						    Char.IsDigit(trimmedCellContent[1]) && Char.IsDigit(trimmedCellContent[2]))
							{ corpsNumbers.Add(5); }
							else
							{ corpsNumbers.Add(null); }							
						}
					}
				}

				for(int col = getColumnNumber(typesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);												
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfTypes.Add(i); }						
					else
					{ types.Add(cellContent); }
				}
				
				for(int col = getColumnNumber(departmentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);												
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfDepartments.Add(i); }						
					else
					{ departments.Add(cellContent); }
				}

				for(int col = getColumnNumber(notUsedColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);												
					if(string.IsNullOrEmpty(cellContent))
					{ notUsed.Add(false); }						
					else
					{ notUsed.Add(true); }
				}

				for(int col = getColumnNumber(placesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);												
					if(string.IsNullOrEmpty(cellContent))
					{ places.Add(0); }						
					else
					{ places.Add(Convert.ToInt32(cellContent)); }
				}				
				close();
				
			}
			catch (Exception ex)
		    {
				reading = false;
		    	MessageBox.Show("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
		    }	
		}

		public override void EvaluateData()
		{
			if(reading)
			{
				const string path = @"E:\BACHELORS WORK\TIMETABLE\DataCollectionApp\BugsReport.txt";
				
				if(missingValuesOfNames.Count != 0 || missingValuesOfTypes.Count !=0 ||
				  missingValuesOfDepartments.Count != 0 || duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("АУДИТОРІЇ.");
						sw.WriteLine("Файл: " + FileName);
					}
				}
				
				if(missingValuesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви аудиторій в рядках: ");
						foreach (int value in missingValuesOfNames)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}

				if(missingValuesOfTypes.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено типи аудиторій в рядках: ");
						foreach (int value in missingValuesOfTypes)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}				

				if(missingValuesOfDepartments.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви кафедр в рядках: ");
						foreach (int value in missingValuesOfDepartments)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				
				if(duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати назв аудиторій: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
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
					
					const string selectAuditoryTypeID = "SELECT auditory_type_id FROM auditory_type WHERE auditory_type_name = @TYPE";
					const string selectDepartmentID = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT_NAME";										
					const string insertAuditories = "INSERT INTO auditory (department_id, auditory_name, not_used, type_auditory, count_of_places, corps_number) VALUES(@ID, @AUDITORY_NAME, @NOT_USED, @TYPE_ID, @COUNT, @CORPS_NUMBER)";
					
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
	                    mySqlCommand.Parameters.AddWithValue("@CORPS_NUMBER", corpsNumbers[i]);
	                    mySqlCommand.ExecuteNonQuery();	
					}				
					connection.Close();
				}
				
                catch (Exception ex)
	            {
					MessageBox.Show("Виникла помилка під час завантаження даних про аудиторії з файлу " + FileName + " до бази даних!");
	            }								
			}
		}
	}
}