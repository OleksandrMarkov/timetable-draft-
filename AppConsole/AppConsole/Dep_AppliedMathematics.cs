using System;

using System.Collections;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;

namespace AppConsole
{
	public class Dep_AppliedMathematics : ExcelFile
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
		
		const char disciplinesColumn = 'B'; // стовпець, з якого беруться назви дисциплін
		const char groupsColumn = 'C'; // стовпець, з якого беруться скорочені назви груп
		const char typesColumn = 'D'; // стовпець, з якого беруться типи занять 
		const char hoursColumn = 'E'; // стовпець, з якого беруться кількості годин на заняття 
		const char controlColumn = 'G'; // стовпець, з якого беруться типи контролю
		const char teachersColumn = 'H'; // стовпець, з якого беруться ПІБ викладачів
		const char auditoriesColumn = 'I'; // стовпець, з якого беруться запропоновані аудиторії
		
		bool reading = true; // стане false, якщо відбудеться помилка при зчитуванні з Excel-файлу
				
		public Dep_AppliedMathematics(string fileName, int firstRow, int lastRow): base(fileName)
		{
			this.fileName = fileName;
			
			this.firstRow = firstRow;
			this.lastRow = lastRow;
		}
		
		public override void ReadFromExcelFile()
		{
			try
			{
				open(1);
				// назви дисциплін
				for(int col = getColumnNumber(disciplinesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					//Console.WriteLine(cellContent + " " + i);
					disciplines.Add(cellContent);
				}
				
				// назви груп
				for(int col = getColumnNumber(groupsColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					groups.Add(cellContent);
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
					teachers.Add(cellContent);
				}
				
				//  запропоновані аудиторії
				for(int col = getColumnNumber(auditoriesColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);					
					auditories.Add(cellContent);
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
			/*if(reading)
			{
				Console.WriteLine("!!!");
			}
			else
			{
				Console.WriteLine("???");
			}*/
		}
		
		public override void Load()
		{
			if(reading)
			{
				try
				{
					MySqlConnection connection = DBUtils.GetDBConnection();
					MySqlCommand mySqlCommand;
					
					const string selectDepartmentID = "SELECT department_id FROM department WHERE short_name = @DEPARTMENT";
					const string selectDisciplineID = "SELECT discipline_id FROM discipline WHERE full_name = @DISCIPLINE";
					
					const string insertLessons = "INSERT INTO lesson (discipline_id, type, countOfHours, control, department_id) "
					+ "VALUES (@DISCIPLINE_ID, @TYPE, @HOURS, @CONTROL, @DEPARTMENT_ID)";
					
					const string selectLessonID = "SELECT lesson_id FROM lesson ORDER BY lesson_id DESC LIMIT 1"; // останнє значення id в Lesson 
								
					const string selectTeacherID = "SELECT teacher_id FROM teacher WHERE full_name = @TEACHER";
					
					const string selectAuditoryID = "SELECT auditory_id FROM auditory WHERE auditory_name = @AUDITORY";
					
					const string insertLesson_teacher = "INSERT INTO lesson_teacher (lesson_id, teacher_id) "
					+ "VALUES (@LESSON_ID, @TEACHER_ID)";
					
					const string insertLesson_auditory = "INSERT INTO lesson_auditory (lesson_id, auditory_id) "
					+ "VALUES (@LESSON_ID, @AUDITORY_ID)";
					
					connection.Open();
					
					mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
					mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", "ЕМ");
					mySqlCommand.ExecuteNonQuery();
						
					int departmentID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
					
					for(int i = 0; i < disciplines.Count; i++)
					{
						mySqlCommand = new MySqlCommand(selectDisciplineID, connection);
						mySqlCommand.Parameters.AddWithValue("@DISCIPLINE", disciplines[i]);
						mySqlCommand.ExecuteNonQuery();
						
						int disciplineID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
						
						//Console.WriteLine(i + " " + disciplines[i] + "\t" + disciplineID);
						
						// вставка в Lesson
						mySqlCommand = new MySqlCommand(insertLessons, connection);
						mySqlCommand.Parameters.AddWithValue("@DISCIPLINE_ID", disciplineID);
						mySqlCommand.Parameters.AddWithValue("@TYPE", lessonsType[i]);
						mySqlCommand.Parameters.AddWithValue("@HOURS", hours[i]);
						mySqlCommand.Parameters.AddWithValue("@CONTROL", lessonsControl[i]);
						mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_ID", departmentID);
						mySqlCommand.ExecuteNonQuery();
						
						mySqlCommand = new MySqlCommand(selectLessonID, connection);
						mySqlCommand.ExecuteNonQuery();
						
						int lessonID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());						
						
						string suggestedAuditories = auditories[i].ToString();
						
						// якщо є запропоновані аудиторії, вони завантажуються в БД (т-ця Lesson_auditory)
						if(suggestedAuditories.Length != 0)
						{
							suggestedAuditories = suggestedAuditories.TrimEnd(new char [] {',', ';'});
							suggestedAuditories = suggestedAuditories.Replace(" ", "");
						
							string [] separatedAuditories = suggestedAuditories.Split(new char[] {',', ';'});
							for (int j = 0; j < separatedAuditories.Length; j++)
							{
								mySqlCommand = new MySqlCommand(selectAuditoryID, connection);
								mySqlCommand.Parameters.AddWithValue("@AUDITORY", separatedAuditories[j]);
								mySqlCommand.ExecuteNonQuery();
						
								int auditoryID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
								
								//Console.WriteLine(auditoryID + "\t" + separatedAuditories[j]);
								
								mySqlCommand = new MySqlCommand(insertLesson_auditory, connection);
								mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
								mySqlCommand.Parameters.AddWithValue("@AUDITORY_ID", auditoryID);
								mySqlCommand.ExecuteNonQuery();
							}	
						}

						string teachersRecord = teachers[i].ToString();
						teachersRecord = teachersRecord.TrimEnd(new char [] {',', ';'});
						//teachersRecord = teachersRecord.Replace(" ", ""); Пробіли є в ПІБ викладачів, вони не видаляються
						//Console.WriteLine(teachersRecord);
						string [] separatedTeachers = teachersRecord.Split(new char[] {',', ';'});
						
						// вставка в Lesson_teacher
						for(int j = 0; j < separatedTeachers.Length; j++)
						{
							separatedTeachers[j] = separatedTeachers[j].Trim(); // прибираються можливі зайві пробіли
							mySqlCommand = new MySqlCommand(selectTeacherID, connection);
							mySqlCommand.Parameters.AddWithValue("@TEACHER", separatedTeachers[j]);
							mySqlCommand.ExecuteNonQuery();
							
							int teacherID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());							
							//Console.WriteLine(j + "\t" + teacherID + "\t" + separatedTeachers[j]);
							
							mySqlCommand = new MySqlCommand(insertLesson_teacher, connection);
							mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
							mySqlCommand.Parameters.AddWithValue("@TEACHER_ID", teacherID);
							mySqlCommand.ExecuteNonQuery();
						}
						
						/*const string insertStudy_group = "INSERT INTO study_group (department_id, full_name, study_group_code) "
						+ "VALUES (@DEPARTMENT_ID, @NAME, @CODE)";
						
						// запис груп в т-цю БД Study_Group
						string groupsInCell = groups[i].ToString();
						if(groupsInCell.Length != 0)
						{
							groupsInCell = groupsInCell.Trim();
							groupsInCell = groupsInCell.TrimEnd(new char [] {',', ';'});
							groupsInCell = groupsInCell.Replace(" ", "");
							groupsInCell = groupsInCell.Replace("\n", "");
							
							//Console.WriteLine(groupsInCell + "!");
							string [] separatedGroups = groupsInCell.Split(new char[] {',', ';'});
							for (int j = 0; j < separatedGroups.Length; j++)
							{
								//коди груп
								int hyphen = separatedGroups[j].IndexOf('-');
								string code = separatedGroups[j].Substring(0, hyphen + 2);
								
								//Console.WriteLine(separatedGroups[j]);
								mySqlCommand = new MySqlCommand(insertStudy_group, connection);
								mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_ID", departmentID);
								mySqlCommand.Parameters.AddWithValue("@NAME", separatedGroups[j]);
								mySqlCommand.Parameters.AddWithValue("@CODE", code);
								mySqlCommand.ExecuteNonQuery();				
							}
							//Console.WriteLine();
						}*/			
					}
					
					/*const string createTemporaryTable = "CREATE TEMPORARY TABLE study_group2 AS (SELECT * FROM study_group GROUP BY full_name)";
	                mySqlCommand = new MySqlCommand(createTemporaryTable, connection);
	                mySqlCommand.ExecuteNonQuery();
	                    	
	                const string deleteTrash = "DELETE FROM study_group WHERE study_group.study_group_id NOT IN (SELECT study_group2.study_group_id FROM study_group2)";
	                mySqlCommand = new MySqlCommand(deleteTrash, connection);
	                mySqlCommand.ExecuteNonQuery();
	               
	                const string dropTemporaryTable = "DROP TABLE study_group2";
	                mySqlCommand = new MySqlCommand(dropTemporaryTable, connection);
	                mySqlCommand.ExecuteNonQuery();	*/				
					
					connection.Close();
					Console.WriteLine("AppliedMathematics Department is loaded!");
				}
				catch(Exception ex)
				{
					Console.WriteLine("Виникла помилка під час запису з файлу " + FileName + " до бази даних!" + "\n" + ex.Message);
				}
			}
		}
	}
}
