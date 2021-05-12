﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Windows;

namespace DataCollectionApp
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
		int row = 1;
		
		const char fullNamesColumn = 'A';
		const char shortNamesColumn = 'B';
		const char facultyCodesColumn = 'C';
		
		bool reading = true;
		
		public Departments(string fileName): base(fileName)
		{ this.fileName = fileName; }
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				for(int col = getColumnNumber(fullNamesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					
					if(full_names.Contains(cellContent))
					{ duplicatesOfFullNames.Add(i, cellContent); }
					else
					{ full_names.Add(cellContent); }
					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfFullNames.Add(i); }
				}
				
				for(int col = getColumnNumber(shortNamesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					if(short_names.Contains(cellContent))
					{ duplicatesOfShortNames.Add(i, cellContent); }
					else
					{ short_names.Add(cellContent); }
					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfShortNames.Add(i); }				
				}
				
				for(int col = getColumnNumber(facultyCodesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfFacultyCodes.Add(i); }
					else
					{ faculty_codes.Add(cellContent); }					
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
				if (missingValuesOfFullNames.Count !=0 || missingValuesOfShortNames.Count != 0
				   || missingValuesOfFacultyCodes.Count != 0 || duplicatesOfFullNames.Count != 0
				  || duplicatesOfShortNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("КАФЕДРИ.");
						sw.WriteLine("Файл: " + FileName);
					}
				}
				if(missingValuesOfFullNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви кафедр в рядках: ");
						foreach (int value in missingValuesOfFullNames)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}			
				if(missingValuesOfShortNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено скорочені назви кафедр в рядках: ");
						foreach (int value in missingValuesOfShortNames)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}			
				if(missingValuesOfFacultyCodes.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено коди факультетів в рядках: ");
						foreach (int value in missingValuesOfFacultyCodes)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}		
				if(duplicatesOfFullNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати назв кафедр: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfFullNames)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}
				}		
				if(duplicatesOfShortNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати скорочених назв кафедр: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfShortNames)
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
					MessageBox.Show("Виникла помилка під час завантаження даних про кафедри з файлу " + FileName + " до бази даних!");
				}
			}
		}	
	}
}