using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class Dep_MachineParts:ExcelFile
	{
		int firstRow; // рядок, з якого починаються записи даних у файлі
		int lastRow; // рядок, на якому закінчуються записи даних у файлі
		
		ArrayList disciplines = new ArrayList();
		ArrayList groups = new ArrayList();
		ArrayList lessonsType = new ArrayList();
		ArrayList hours = new ArrayList();
		ArrayList lessonsControl = new ArrayList();
		ArrayList teachers = new ArrayList();
		ArrayList auditories = new ArrayList();
		
		//ArrayList missingValuesOfDisciplines = new ArrayList();
		
		const char disciplinesColumn = 'B'; // стовпець, з якого беруться назви дисциплін
		const char groupsColumn = 'C'; // стовпець, з якого беруться скорочені назви груп
		const char typesColumn = 'D'; // стовпець, з якого беруться типи занять 
		const char hoursColumn = 'E'; // стовпець, з якого беруться кількості годин на заняття 
		const char controlColumn = 'G'; // стовпець, з якого беруться типи контролю
		const char teachersColumn = 'H'; // стовпець, з якого беруться ПІБ викладачів
		const char auditoriesColumn = 'I'; // стовпець, з якого беруться запропоновані аудиторії
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
				
		public Dep_MachineParts(string fileName, int firstRow, int lastRow): base(fileName)
		{
			this.fileName = fileName;
			
			this.firstRow = firstRow;
			this.lastRow = lastRow;
		}
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open();
				
				// назви дисциплін
				for(int col = getColumnNumber(disciplinesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					//Console.WriteLine(cellContent + " " + i);
					disciplines.Add(cellContent);
				}
				
				//  назви груп
				for(int col = getColumnNumber(groupsColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					
					// в конце перечня груп может стоять случайно забытая запятая, которая все ломает
					cellContent = cellContent.TrimEnd(',');
					
					// прибираються пробіли
					cellContent = cellContent.Replace(" ", "");
					
					// групи розділені ';' або ','
					string [] groupsInCell = cellContent.Split(new char[] {',', ';'});
					
					/*foreach (string g in groupsInCell)
					{
						Console.Write(g + "! ");
					}
					Console.Write(groupsInCell.Length);
					Console.WriteLine();*/
					
					groups.Add(groupsInCell);
				}
				
				//типи занять
				for(int col = getColumnNumber(typesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					lessonsType.Add(cellContent);
				}
				
				//кількості годин
				for(int col = getColumnNumber(hoursColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					int h = Convert.ToInt32(cellContent);
					hours.Add(h);
				}
				
				//типи контролю
				for(int col = getColumnNumber(controlColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					lessonsControl.Add(cellContent);
				}
				
				
				//  викладачі
				for(int col = getColumnNumber(teachersColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					
					cellContent = cellContent.TrimEnd(';');
					
					cellContent = cellContent.Replace(" ", "");
					
					string [] teachersInCell = cellContent.Split(new char[] {',', ';'});
					
					teachers.Add(teachersInCell);
				}
				
				//  запропоновані аудиторії
				for(int col = getColumnNumber(auditoriesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					
					cellContent = cellContent.TrimEnd(';');
					
					cellContent = cellContent.Replace(" ", "");
					
					string [] auditoriesInCell = cellContent.Split(new char[] {',', ';'});
					
					auditories.Add(auditoriesInCell);
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
					
					Console.WriteLine("MachineParts Department is loaded!");
				}
				catch(Exception ex)
				{
					Console.WriteLine("Виникла помилка під час запису з файлу " + FileName + " до бази даних!");
				}
			}
		}
		
	}
}
