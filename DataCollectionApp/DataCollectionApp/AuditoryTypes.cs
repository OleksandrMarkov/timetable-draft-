﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace DataCollectionApp
{
	public class AuditoryTypes : ExcelFile
	{			
		ArrayList records = new ArrayList();
		ArrayList missingValues = new ArrayList();
		Dictionary <int, string> duplicates = new Dictionary<int, string>();
		int row = 1;
		const char column = 'A';
		bool reading = true;

		DbOperations dbo = new DbOperations();				
		
		public AuditoryTypes(string fileName): base(fileName)
		{ this.fileName = fileName; }
				
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);				
				for(int col = getColumnNumber(column), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);					
					if(records.Contains(cellContent))
					{ duplicates.Add(i, cellContent); }
					else
					{ records.Add(cellContent); }
					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValues.Add(i); }
				}				
				close();
			}
			catch(Exception ex)
			{
				reading = false;
				MessageBox.Show("Помилка при отриманні даних з файлу " + FileName + ". " + ex.Message);
			}			
		}
		
		public override void EvaluateData()
		{
			if (reading)
			{
				const string path = @"E:\BACHELORS WORK\TIMETABLE\DataCollectionApp\BugsReport.txt";
				if(missingValues.Count !=0 || duplicates.Count !=0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("ТИПИ АУДИТОРІЙ.");
						sw.WriteLine("Файл: " + FileName);
					}	
				}
				
				if (missingValues.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є пропуски в рядках: ");
						foreach (int value in missingValues)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				
				if(duplicates.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати: ");
						foreach (KeyValuePair <int, string> duplicate in duplicates)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}					
				}
				
				bool noSensetoReload = true;				
				ArrayList auditoryTypesInDB = new ArrayList();
				auditoryTypesInDB = dbo.getAuditory_types();

				foreach (string record in records)
				{
					if(!auditoryTypesInDB.Contains(record))
					{
						noSensetoReload = false;
						break;
					}
				}
				if (noSensetoReload)
				{
					reading = false;
					MessageBox.Show("Дані про типи аудиторій вже містяться в базі даних!");
				}			
			}
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{
					foreach (string record in records)
					{
						dbo.insertAuditory_type(record);
					}
					MessageBox.Show("Дані про типи аудиторій завантажено до бази даних!");					
				}
				catch(Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження даних про типи аудиторій з файлу " + FileName + " до бази даних!");
				}	
			}			
		}
	}
}