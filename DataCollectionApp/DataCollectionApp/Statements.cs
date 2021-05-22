using System;
using System.Collections;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Windows;

using System.IO;

namespace DataCollectionApp
{
	public class Statements : ExcelFile
	{
		int firstRow;
		public int lastRow;
		
		ArrayList disciplines = new ArrayList();
		ArrayList groups = new ArrayList();
		ArrayList lessonsType = new ArrayList();
		ArrayList hours = new ArrayList();
		ArrayList lessonsControl = new ArrayList();
		ArrayList teachers = new ArrayList();
		ArrayList auditories = new ArrayList();
		ArrayList days = new ArrayList();
		
		const char disciplinesColumn = 'B';
		const char groupsColumn = 'C';
		const char typesColumn = 'D';		
		const char hoursColumn = 'F';
		const char controlColumn = 'G';
		const char teachersColumn = 'H';
		const char auditoriesColumn = 'I';
		const char daysColumn = 'J';
		
		int sheetNumber;
		string departmentShortName;
		
		bool reading = true;
				
		public Statements(string fileName, int firstRow, int sheetNumber, string departmentShortName): base(fileName)
		{
			this.fileName = fileName;
			this.firstRow = firstRow;
			this.sheetNumber = sheetNumber;
			this.departmentShortName = departmentShortName;
			this.lastRow = getLastRow(fileName, sheetNumber, firstRow);
		}
		
		DbOperations dbo = new DbOperations();
		CellContentCorrection ccc = new CellContentCorrection();
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open(sheetNumber);
				for(int col = getColumnNumber(disciplinesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					disciplines.Add(cellContent);
				}
				for(int col = getColumnNumber(groupsColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					groups.Add(cellContent);
				}
				for(int col = getColumnNumber(typesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					lessonsType.Add(cellContent);
				}
				for(int col = getColumnNumber(hoursColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					if (string.IsNullOrEmpty(cellContent))
					{
						hours.Add(null);
					}
					else
					{
						char comma = ',';
						int index = cellContent.IndexOf(comma);
					
						if(index != -1)
						{ cellContent = cellContent.Substring(0, cellContent.Length - index - 1); }			
						double h = Convert.ToDouble(cellContent);
						hours.Add(h);	
					}
				}
				for(int col = getColumnNumber(controlColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					lessonsControl.Add(cellContent);
				}
				for(int col = getColumnNumber(teachersColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					teachers.Add(cellContent);
				}
				for(int col = getColumnNumber(auditoriesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);					
					auditories.Add(cellContent);
				}
				for(int col = getColumnNumber(daysColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					days.Add(cellContent);
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
			int departmentID = dbo.getDepartmentID(departmentShortName);
			
			if (SheetsCount == 1)
			{
				dbo.deleteExcessData(departmentID);	
			}
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{	
					const string path = @"E:\BACHELORS WORK\TIMETABLE\DataCollectionApp\BugsReport.txt";
					
					int departmentID = dbo.getDepartmentID(departmentShortName);
					
					for(int i = 0; i < disciplines.Count; i++)
					{
						int disciplineID = dbo.getDisciplineID(disciplines[i].ToString());
						
						dbo.insertLesson(disciplineID, lessonsType[i].ToString(), Convert.ToInt32(hours[i]),
						lessonsControl[i].ToString(), departmentID);
						
						int lastLessonID = dbo.getLastLessonID();
												
						string suggestedAuditories = auditories[i].ToString();
						
						if(suggestedAuditories.Length != 0)
						{					
							suggestedAuditories = ccc.correctAuditories(suggestedAuditories);
							
							string [] separatedAuditories = suggestedAuditories.Split(new char[] {',', ';'});
							
							for (int j = 0; j < separatedAuditories.Length; j++)
							{
								ArrayList auditoriesInDB  = dbo.getAuditoryNames();
								
								if(auditoriesInDB.Contains(separatedAuditories[j]) == false)
								{
									dbo.insertAuditory(departmentID, separatedAuditories[j]);
									using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
									{
										sw.WriteLine("{0:g}", DateTime.Now);
										sw.Write("Аудиторію \"" + separatedAuditories[j] + "\" кафедри \"" + departmentShortName +
										         "\" завантажено до бази даних з файлу \"" + FileName + "\"");
										sw.WriteLine();
									}
								}
								
								int auditoryID = dbo.getAuditoryID(separatedAuditories[j]);
								dbo.insertLesson_Auditory(lastLessonID, auditoryID);
							}	
						}

						string teachersRecord = teachers[i].ToString();
						teachersRecord = ccc.correctTeachers(teachersRecord);
						
						string [] separatedTeachers = teachersRecord.Split(new char[] {',', ';'});
			
						for(int j = 0; j < separatedTeachers.Length; j++)
						{
							separatedTeachers[j] = separatedTeachers[j].Trim();
							
							int teacherID = dbo.getTeacherID(separatedTeachers[j]);
							
							dbo.insertLesson_Teacher(lastLessonID, teacherID);
						}
						
						string groupsInCell = groups[i].ToString();
						if(groupsInCell.Length != 0)
						{
							groupsInCell = ccc.correctGroups(groupsInCell);
							
							string [] separatedGroups = groupsInCell.Split(new char[] {',', ';'});
							
							ArrayList studyGroups = dbo.getStudy_groups();
								
							for(int j = 0; j < separatedGroups.Length; j++)
							{
								if (studyGroups.Contains(separatedGroups[j]) == false)
								{
									GroupCode gc = new GroupCode();
									string code = gc.getGroupCode(separatedGroups[j]);
									
									dbo.insertStudy_group(departmentID, separatedGroups[j], code);
									using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
									{
										sw.WriteLine("{0:g}", DateTime.Now);
										sw.Write("Учбову групу \"" + separatedGroups[j] + "\" кафедри \"" + departmentShortName +
										         "\" завантажено до бази даних з файлу \"" + FileName + "\"");
										sw.WriteLine();
									}									
								}
								int groupID = dbo.getStudy_groupID(separatedGroups[j]);
								dbo.insertLesson_group(lastLessonID, groupID);
							}
						}
						
						string daysInCell = days[i].ToString();
						if(daysInCell.Length != 0)
						{
							daysInCell = ccc.correctTime(daysInCell);
							string [] separatedDays = daysInCell.Split(new char[] {',', ';'});
							for(int j = 0; j < separatedDays.Length; j++)
							{
								dbo.insertLesson_time(lastLessonID, separatedDays[j]);
							}
						}
					}		
				}
				catch(Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження відомостей доручень до бази даних!" + "\n" + ex.Message);
				}
			}
		}
	}
}