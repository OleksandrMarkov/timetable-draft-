﻿using System;
using System.Collections;
using System.Collections.Generic;

using System.IO;
using System.Windows;

namespace DataCollectionApp
{
	public class Teachers : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList sex = new ArrayList();
		ArrayList posts = new ArrayList();
		ArrayList statuses = new ArrayList();
		ArrayList departments = new ArrayList();
		
		const char departmentsColumn = 'A';
		const char namesColumn = 'J';
		const char sexColumn = 'K';
		const char postsColumn = 'L';
		const char statusesColumn = 'M';
		
		int row = 8;
		
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfPosts = new ArrayList();
		ArrayList missingValuesOfDepartments = new ArrayList();	
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		Dictionary <int, string> wrongValuesOfSex = new Dictionary<int, string>();
		
		bool reading = true;
		
		
		DbOperations dbo = new DbOperations();
		
		
		public Teachers(string fileName): base(fileName)
		{ this.fileName = fileName; }
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				for(int col = getColumnNumber(departmentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfDepartments.Add(i); }			
					else
					{ departments.Add(cellContent); }
				}	
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);	
					if(names.Contains(cellContent))
					{ duplicatesOfNames.Add(i, cellContent); }
					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfNames.Add(i); }
					names.Add(cellContent);
				}	
				for(int col = getColumnNumber(sexColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					sex.Add(cellContent);	
					if ( cellContent != "м" && cellContent != "ж")
                    { wrongValuesOfSex.Add(i, cellContent); }
				}
				for(int col = getColumnNumber(postsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);	
					posts.Add(cellContent);
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfPosts.Add(i); }
				}
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
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				if (missingValuesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено ПІБ викладачів в рядках: ");
						foreach (int value in missingValuesOfNames)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				if (missingValuesOfPosts.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено посади викладачів в рядках: ");
						foreach (int value in missingValuesOfPosts)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				if(duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати ПІБ викладачів: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}
				}
				if(wrongValuesOfSex.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є некоректні значення статі викладачів: ");
						foreach (KeyValuePair<int, string> wrongValue in wrongValuesOfSex)
						{ sw.WriteLine("В рядку номер " + wrongValue.Key + ": " + wrongValue.Value); }
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
					for(int i = 0; i < rowsCount - row + 1; i++)
					{
						int departmentID = dbo.getDepartmentIDbyFullName(departments[i].ToString());
						dbo.insertTeacher(departmentID, names[i].ToString(), sex[i].ToString(),
						posts[i].ToString(), statuses[i].ToString());
					}
					dbo.correctTeacherTable();	
				}
				catch (Exception ex)
	            {
					MessageBox.Show("Виникла помилка під час завантаження даних про викладачів з файлу " + FileName + " до бази даних!" + "\n " + ex.Message);
				}
			}
		}
		
	}
}