using System;
using System.Collections;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Windows;

namespace DataCollectionApp
{
	public class Dep_AccountingAndAudit : ExcelFile
	{
		int firstRow;
		int lastRow;
		ArrayList disciplines = new ArrayList();
		ArrayList groups = new ArrayList();
		ArrayList lessonsType = new ArrayList();
		ArrayList hours = new ArrayList();
		ArrayList lessonsControl = new ArrayList();
		ArrayList teachers = new ArrayList();
		ArrayList auditories = new ArrayList();
		
		const char disciplinesColumn = 'B';
		const char groupsColumn = 'C';
		const char typesColumn = 'D';
		const char hoursColumn = 'E';
		const char controlColumn = 'G';
		const char teachersColumn = 'H';
		const char auditoriesColumn = 'I';
		
		bool reading = true;	
		int sheetNumber;
		
		public Dep_AccountingAndAudit(string fileName, int firstRow, int lastRow, int sheetNumber): base(fileName)
		{
			this.fileName = fileName;
			this.firstRow = firstRow;
			this.lastRow = lastRow;
			this.sheetNumber = sheetNumber;
		}
		
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
					cellContent = cellContent.Trim();
					lessonsType.Add(cellContent);
				}				
				for(int col = getColumnNumber(hoursColumn), i = firstRow; i <= lastRow; i++)
				{
					cellContent = getCellContent(i, col);
					int h = Convert.ToInt32(cellContent);
					hours.Add(h);
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
				close();
			}			
			catch (Exception ex)
            {
				reading = false;
            	MessageBox.Show("Помилка при отриманні даних з файлу " + FileName + " " + ex.Message);
            }
		}
		
		public override void EvaluateData(){}
		
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
					const string selectLessonID = "SELECT lesson_id FROM lesson ORDER BY lesson_id DESC LIMIT 1";						
					const string selectTeacherID = "SELECT teacher_id FROM teacher WHERE full_name = @TEACHER";		
					const string selectAuditoryID = "SELECT auditory_id FROM auditory WHERE auditory_name = @AUDITORY";		
					const string insertLesson_teacher = "INSERT INTO lesson_teacher (lesson_id, teacher_id) "
					+ "VALUES (@LESSON_ID, @TEACHER_ID)";		
					const string insertLesson_auditory = "INSERT INTO lesson_auditory (lesson_id, auditory_id) "
					+ "VALUES (@LESSON_ID, @AUDITORY_ID)";
					
					connection.Open();
					
					mySqlCommand = new MySqlCommand(selectDepartmentID, connection);
					mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", "ОіО");
					mySqlCommand.ExecuteNonQuery();
						
					int departmentID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
					
					for(int i = 0; i < disciplines.Count; i++)
					{
						mySqlCommand = new MySqlCommand(selectDisciplineID, connection);
						mySqlCommand.Parameters.AddWithValue("@DISCIPLINE", disciplines[i]);
						mySqlCommand.ExecuteNonQuery();
						
						int disciplineID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
						
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
															
								mySqlCommand = new MySqlCommand(insertLesson_auditory, connection);
								mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
								mySqlCommand.Parameters.AddWithValue("@AUDITORY_ID", auditoryID);
								mySqlCommand.ExecuteNonQuery();
							}	
						}

						string teachersRecord = teachers[i].ToString();
						teachersRecord = teachersRecord.TrimEnd(new char [] {',', ';'});
						string [] separatedTeachers = teachersRecord.Split(new char[] {',', ';'});
						
						for(int j = 0; j < separatedTeachers.Length; j++)
						{
							separatedTeachers[j] = separatedTeachers[j].Trim();
							mySqlCommand = new MySqlCommand(selectTeacherID, connection);
							mySqlCommand.Parameters.AddWithValue("@TEACHER", separatedTeachers[j]);
							mySqlCommand.ExecuteNonQuery();
							
							int teacherID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());							
							
							mySqlCommand = new MySqlCommand(insertLesson_teacher, connection);
							mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
							mySqlCommand.Parameters.AddWithValue("@TEACHER_ID", teacherID);
							mySqlCommand.ExecuteNonQuery();
						}
						
						const string selectStudyGroups = "SELECT full_name FROM study_group";
						const string insertLesson_group = "INSERT INTO lesson_group (lesson_id, group_id) "
						+ "VALUES (@LESSON_ID, @GROUP_ID)";
						const string insertStudy_group = "INSERT INTO study_group (department_id, full_name, study_group_code) "
						+ "VALUES (@DEPARTMENT_ID, @NAME, @CODE)";					
						const string selectGroupID = "SELECT study_group_id FROM study_group WHERE full_name = @GROUP";
						
						string groupsInCell = groups[i].ToString();
						if(groupsInCell.Length != 0)
						{
							groupsInCell = groupsInCell.Trim();
							groupsInCell = groupsInCell.TrimEnd(new char [] {',', ';'});
							groupsInCell = groupsInCell.Replace(" ", "");
							groupsInCell = groupsInCell.Replace("\n", "");
							
							string [] separatedGroups = groupsInCell.Split(new char[] {',', ';'});
							
							ArrayList studyGroupsFromDB = new ArrayList();
							mySqlCommand = new MySqlCommand(selectStudyGroups, connection);
							using (MySqlConnection connection2 = DBUtils.GetDBConnection())
							{
								using(MySqlDataReader dataReader = mySqlCommand.ExecuteReader())
								{
									while(dataReader.Read())
									{ studyGroupsFromDB.Add(dataReader[0].ToString()); }
								}
							}
							
							for(int j = 0; j < separatedGroups.Length; j++)
							{
								if (studyGroupsFromDB.Contains(separatedGroups[j]) == false)
								{
									int hyphen = separatedGroups[j].IndexOf('-');
									string code = separatedGroups[j].Substring(0, hyphen + 2);
									
									mySqlCommand = new MySqlCommand(insertStudy_group, connection);
									mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_ID", departmentID);
									mySqlCommand.Parameters.AddWithValue("@NAME", separatedGroups[j]);
									mySqlCommand.Parameters.AddWithValue("@CODE", code);
									mySqlCommand.ExecuteNonQuery();
								}
								mySqlCommand = new MySqlCommand(selectGroupID, connection);
								mySqlCommand.Parameters.AddWithValue("@GROUP", separatedGroups[j]);
								mySqlCommand.ExecuteNonQuery();
								int groupID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
								
								mySqlCommand = new MySqlCommand(insertLesson_group, connection);
								mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
								mySqlCommand.Parameters.AddWithValue("@GROUP_ID", groupID);
								mySqlCommand.ExecuteNonQuery();
							}
						}
					}					
					connection.Close();
				}
				catch(Exception ex)
				{
					MessageBox.Show("Виникла помилка під час завантаження відомостей доручень до бази даних!" + "\n" + ex.Message);
				}
			}
		}
	}
}