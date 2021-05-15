using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace DataCollectionApp
{
	public class StudyGroups : ExcelFile
	{		
		public StudyGroups(string fileName): base(fileName)
		{ this.fileName = fileName; }
		
		ArrayList departments = new ArrayList();
		ArrayList names = new ArrayList();
		ArrayList courseNumbers = new ArrayList();
		ArrayList countOfStudents = new ArrayList();		
		ArrayList codes = new ArrayList();
		ArrayList missingValuesOfDepartments = new ArrayList();
		ArrayList missingValuesOfNames = new ArrayList();			
		Dictionary <int, string> duplicatesOfNames = new Dictionary<int, string>();
		
		const char departmentsColumn = 'A';
		const char namesColumn = 'B';
		const char countOfStudentsColumn = 'C';
		const char courseNumbersColumn = 'D';
		
		bool reading = true;
		int row = 1;
		
		DbOperations dbo = new DbOperations();
		
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
					{
						cellContent = cellContent.Replace("*", "");
						departments.Add(cellContent);
					}
				}
				for(int col = getColumnNumber(namesColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);						
					if(string.IsNullOrEmpty(cellContent))
					{ missingValuesOfNames.Add(i); }
					else
					{
						if(names.Contains(cellContent))
						{ duplicatesOfNames.Add(i, cellContent); }
						
						names.Add(cellContent);
						int hyphen = cellContent.IndexOf('-');
						codes.Add(cellContent.Substring(0, hyphen + 2));
					}
				}			
				for(int col = getColumnNumber(courseNumbersColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);
					courseNumbers.Add(cellContent);
				}
				for(int col = getColumnNumber(countOfStudentsColumn), i = row; i <= rowsCount; i++)
				{
					cellContent = getCellContent(i, col);					
					if(string.IsNullOrEmpty(cellContent))
					{ countOfStudents.Add(null); }					
					else
					{ countOfStudents.Add(Convert.ToInt32(cellContent)); }
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
				if(missingValuesOfNames.Count != 0 || missingValuesOfDepartments.Count != 0
				  || duplicatesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("{0:g}", DateTime.Now);
						sw.WriteLine("УЧБОВІ ГРУПИ.");
						sw.WriteLine("Файл: " + FileName);
					}
				}
				if(missingValuesOfNames.Count != 0)
				{
					using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Пропущено назви груп в рядках: ");
						foreach (int value in missingValuesOfNames)
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
                if (duplicatesOfNames.Count != 0)
                {
                	using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
					{
						sw.WriteLine("Є дублікати назв груп: ");
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
					for (int i = 0; i < rowsCount; i++)
		            {
						int departmentID = dbo.getDepartmentIDbyFullName(departments[i].ToString());
						dbo.insertFullDataStudy_group(departmentID, codes[i].ToString(), names[i].ToString(),
						courseNumbers[i].ToString(), Convert.ToInt32(countOfStudents[i]));
					}                
                 }
				catch (Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження даних про учбові групи з файлу " + FileName + " до бази даних!");
				}				
			}
		}
	}
}