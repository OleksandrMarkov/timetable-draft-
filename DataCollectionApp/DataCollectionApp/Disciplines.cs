using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace DataCollectionApp
{
	public class Disciplines : ExcelFile
	{
		ArrayList records = new ArrayList();
		ArrayList missingValues = new ArrayList();
		Dictionary <int, string> duplicates = new Dictionary<int, string>();
		int row = 2;
		const char column = 'G';
		bool reading = true;
		
		DbOperations dbo = new DbOperations();
		
		public Disciplines(string fileName): base(fileName)
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
						sw.WriteLine("ДИСЦИПЛІНИ.");
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
						foreach (KeyValuePair<int, string> duplicate in duplicates)
						{ sw.WriteLine("В рядку номер " + duplicate.Key + ": " + duplicate.Value); }
						sw.WriteLine();
					}					
				}

				bool noSensetoReload = true;
				ArrayList disciplinesInDB = dbo.getDisciplines();
								
				foreach (string record in records)
				{
					if(!disciplinesInDB.Contains(record))
					{
						noSensetoReload = false;
						break;
					}
				}
				if (noSensetoReload)
				{
					reading = false;
					MessageBox.Show("Дані про дисципліни вже містяться в базі даних!");
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
						dbo.insertDiscipline(record);
					}
					MessageBox.Show("Дані про дисципліни завантажено до бази даних!");					
				}
				catch(Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження даних про дисципліни з файлу " + FileName + " до бази даних!");
				}	
			}			
		}
	}
}