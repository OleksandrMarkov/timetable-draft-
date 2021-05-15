using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace DataCollectionApp
{
	public class Faculties : ExcelFile
	{
		ArrayList names = new ArrayList();
		ArrayList codes = new ArrayList();	
		ArrayList missingValuesOfNames = new ArrayList();
		ArrayList missingValuesOfCodes = new ArrayList();	
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		Dictionary <int, string> duplicatesOfCodes = new Dictionary<int, string>();
		int row = 2;
		const char namesColumn = 'B';
		const char codesColumn = 'D';
		bool reading = true;
		
		DbOperations dbo = new DbOperations();
		
		public Faculties(string fileName): base(fileName)
		{ this.fileName = fileName; }
			
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					if(names.Contains(cellContent))
					{ duplicatesOfNames.Add(i, cellContent); }	
					else
					{ names.Add(cellContent); }
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfNames.Add(i); }
				}
				for(int col = getColumnNumber(codesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					if(codes.Contains(cellContent))
					{ duplicatesOfCodes.Add(i, cellContent); }
					else
					{ codes.Add(cellContent); }
					
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfCodes.Add(i); }				
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
				
				if(missingValuesOfNames.Count != 0 || duplicatesOfNames.Count !=0 
				  || missingValuesOfCodes.Count != 0 || duplicatesOfCodes.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("ФАКУЛЬТЕТИ.");
						sw.WriteLine("Файл: " + FileName);
					}
				}
				if(missingValuesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви факультетів в рядках: ");
						foreach (int value in missingValuesOfNames)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}	
				if(duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати назв факультетів: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfNames)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}
				}
				if(missingValuesOfCodes.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено коди факультетів в рядках: ");
						foreach (int value in missingValuesOfCodes)
						{ sw.Write(value + "|"); }
						sw.WriteLine();
					}
				}
				if(duplicatesOfCodes.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати кодів факультетів: ");
						foreach (KeyValuePair<int, string> duplicate in duplicatesOfCodes)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}
				}
	
				bool noSensetoReload = true;
				
				ArrayList facultyNamesInDB = dbo.getFacultyNames();
				ArrayList facultyCodesInDB = dbo.getFacultyCodes();
								
				foreach (string name in names)
				{
					if(!facultyNamesInDB.Contains(name))
					{
						noSensetoReload = false;
						break;
					}
				}
				foreach (string code in codes)
				{
					if(!facultyCodesInDB.Contains(code))
					{
						noSensetoReload = false;
						break;
					}
				}		
				if (noSensetoReload)
				{
					reading = false;
					MessageBox.Show("Дані про факультети вже містяться в базі даних!");
				}				
			}						
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{
					for(int i = 0; i < names.Count; i++)
					{
						dbo.insertFaculty(names[i].ToString(), codes[i].ToString());
					}
					MessageBox.Show("Дані про факультети завантажено до бази даних!");
				}
				catch(Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження даних про факультети з файлу " + FileName + " до бази даних!");
				}			
			}			
		}		
	}
}