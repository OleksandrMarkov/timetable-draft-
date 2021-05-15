﻿using System;
using System.Collections;
using System.Collections.Generic;

using MySql.Data.MySqlClient;

namespace DataCollectionApp
{
	public class DbOperations
	{	
		MySqlConnection mySqlConnection = DBUtils.GetDBConnection();
		MySqlCommand mySqlCommand;
		MySqlDataReader mySqlDataReader;
		
		public Dictionary<int, string> getDepartments()
		{
			Dictionary <int, string> departments = new Dictionary<int, string>();
			mySqlConnection.Open();
			const string getDepartments = "SELECT department_id, full_name FROM department WHERE department_id < 328";
			mySqlCommand = new MySqlCommand (getDepartments, mySqlConnection);
			
			using(mySqlDataReader = mySqlCommand.ExecuteReader())
			{
				while(mySqlDataReader.Read())
				{
					departments.Add(Convert.ToInt32(mySqlDataReader[0]), mySqlDataReader[1].ToString());
				}
			}
			mySqlConnection.Close();		
			return departments;		
		}
		
		public int getDepartmentID(string departmentName)
		{
			const string command = "SELECT department_id FROM department WHERE short_name = @DEPARTMENT";	
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", departmentName);
			mySqlCommand.ExecuteNonQuery();
			int departmentID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return departmentID;
		}
		
		public int getDisciplineID(string discipline)
		{
			const string command = "SELECT discipline_id FROM discipline WHERE full_name = @DISCIPLINE";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@DISCIPLINE", discipline);
			mySqlCommand.ExecuteNonQuery();
			int disciplineID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return disciplineID;
		}
		
		public void insertLesson(int disciplineID, string lessonsType, int hours, string lessonsControl, int departmentID)
		{
			const string command = "INSERT INTO lesson (discipline_id, type, countOfHours, control, department_id) "
					+ "VALUES (@DISCIPLINE_ID, @TYPE, @HOURS, @CONTROL, @DEPARTMENT_ID)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@DISCIPLINE_ID", disciplineID);
			mySqlCommand.Parameters.AddWithValue("@TYPE", lessonsType);
			mySqlCommand.Parameters.AddWithValue("@HOURS", hours);
			mySqlCommand.Parameters.AddWithValue("@CONTROL", lessonsControl);
			mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_ID", departmentID);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public int getLastLessonID()
		{
			const string command = "SELECT lesson_id FROM lesson ORDER BY lesson_id DESC LIMIT 1";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.ExecuteNonQuery();
			int lessonID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());			
			mySqlConnection.Close();
			return lessonID;
		}
		
		public int getAuditoryID(string auditory)
		{
			const string command = "SELECT auditory_id FROM auditory WHERE auditory_name = @AUDITORY";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@AUDITORY", auditory);
			mySqlCommand.ExecuteNonQuery();
			int auditoryID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return auditoryID;
		}
		
		public void insertLesson_Auditory(int lastLessonID, int auditoryID)
		{
			const string command = "INSERT INTO lesson_auditory (lesson_id, auditory_id) "
					+ "VALUES (@LESSON_ID, @AUDITORY_ID)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lastLessonID);
			mySqlCommand.Parameters.AddWithValue("@AUDITORY_ID", auditoryID);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public int getTeacherID(string teacher)
		{
			const string command = "SELECT teacher_id FROM teacher WHERE full_name = @TEACHER";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@TEACHER", teacher);
			mySqlCommand.ExecuteNonQuery();
			int teacherID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return teacherID;
		}
		
		public void insertLesson_Teacher(int lastLessonID, int teacherID)
		{
			const string command = "INSERT INTO lesson_teacher (lesson_id, teacher_id) "
					+ "VALUES (@LESSON_ID, @TEACHER_ID)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lastLessonID);
			mySqlCommand.Parameters.AddWithValue("@TEACHER_ID", teacherID);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public ArrayList getStudy_groups()
		{
			const string command = "SELECT full_name FROM study_group";
			ArrayList studyGroups = new ArrayList();
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlDataReader = mySqlCommand.ExecuteReader();
			while(mySqlDataReader.Read())
			{
				studyGroups.Add(mySqlDataReader[0].ToString());
			}
			mySqlConnection.Close();
			return studyGroups;
		}
		
		public void insertStudy_group(int departmentID, string name, string code)
		{
			const string command = "INSERT INTO study_group (department_id, full_name, study_group_code) "
			+ "VALUES (@DEPARTMENT_ID, @NAME, @CODE)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@DEPARTMENT_ID", departmentID);
			mySqlCommand.Parameters.AddWithValue("@NAME", name);
			mySqlCommand.Parameters.AddWithValue("@CODE", code);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public int getStudy_groupID(string study_group)
		{
			const string command = "SELECT study_group_id FROM study_group WHERE full_name = @GROUP";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@GROUP", study_group);
			mySqlCommand.ExecuteNonQuery();
			int study_groupID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return study_groupID;	
		}
		
		public void insertLesson_group(int lessonID, int groupID)
		{
			const string command = "INSERT INTO lesson_group (lesson_id, group_id) "
			+ "VALUES (@LESSON_ID, @GROUP_ID)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@LESSON_ID", lessonID);
			mySqlCommand.Parameters.AddWithValue("@GROUP_ID", groupID);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public ArrayList getAuditory_types()
		{
			const string command = "SELECT auditory_type_name FROM auditory_type";
			ArrayList auditoryTypes = new ArrayList();
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlDataReader = mySqlCommand.ExecuteReader();
			while(mySqlDataReader.Read())
			{
				auditoryTypes.Add(mySqlDataReader[0].ToString());
			}
			mySqlConnection.Close();
			return auditoryTypes;
		}		
		
		public void insertAuditory_type(string auditory_type)
		{
			const string command = "INSERT INTO auditory_type (auditory_type_name) VALUES (@TYPE)";
			
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@TYPE", auditory_type);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public ArrayList getDisciplines()
		{
			const string command =  "SELECT full_name FROM discipline";
			ArrayList disciplines = new ArrayList();
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlDataReader = mySqlCommand.ExecuteReader();
			while(mySqlDataReader.Read())
			{
				disciplines.Add(mySqlDataReader[0].ToString());
			}
			mySqlConnection.Close();
			return disciplines;
		}

		public void insertDiscipline(string discipline)
		{
			const string command = "INSERT INTO discipline (full_name) VALUES (@FULL_NAME)";
			
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@FULL_NAME", discipline);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public ArrayList getFacultyNames()
		{
			const string command =  "SELECT full_name FROM faculty";
			ArrayList facultyNames = new ArrayList();
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlDataReader = mySqlCommand.ExecuteReader();
			while(mySqlDataReader.Read())
			{
				facultyNames.Add(mySqlDataReader[0].ToString());
			}
			mySqlConnection.Close();
			return facultyNames;
		}		

		public ArrayList getFacultyCodes()
		{
			const string command =  "SELECT faculty_code FROM faculty";
			ArrayList facultyCodes = new ArrayList();
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlDataReader = mySqlCommand.ExecuteReader();
			while(mySqlDataReader.Read())
			{
				facultyCodes.Add(mySqlDataReader[0].ToString());
			}
			mySqlConnection.Close();
			return facultyCodes;
		}

		public void insertFaculty(string name, string code)
		{
			const string command = "INSERT INTO faculty (full_name, faculty_code) VALUES (@FULL_NAME, @CODE)";
			
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@FULL_NAME", name);
			mySqlCommand.Parameters.AddWithValue("@CODE", code);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public int getFacultyID(string code)
		{
			const string command = "SELECT faculty_id FROM faculty WHERE faculty_code = @CODE";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@CODE", code);
			mySqlCommand.ExecuteNonQuery();
			int facultyID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return facultyID;	
		}

		public void insertDepartment(int facultyID, string fullName, string shortName)
		{
			const string command ="INSERT INTO department (faculty_id, full_name, short_name) "
			+ "VALUES (@FACULTY_ID, @FULL_NAME, @SHORT_NAME)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@FACULTY_ID", facultyID);
			mySqlCommand.Parameters.AddWithValue("@FULL_NAME", fullName);
			mySqlCommand.Parameters.AddWithValue("@SHORT_NAME", shortName);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public int getDepartmentIDbyFullName(string departmentName)
		{
			const string command = "SELECT department_id FROM department WHERE full_name = @DEPARTMENT";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@DEPARTMENT", departmentName);
			mySqlCommand.ExecuteNonQuery();
			int departmentID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return departmentID;	
		}


		public void insertTeacher(int departmentID, string name, string sex, string post, string status)
		{
			const string command = "INSERT INTO teacher (department_id, full_name, sex, post, status) VALUES(@ID, @NAME, @SEX, @POST, @STATUS)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
			mySqlCommand.Parameters.AddWithValue("@NAME", name);
			mySqlCommand.Parameters.AddWithValue("@SEX", sex);
			mySqlCommand.Parameters.AddWithValue("@POST", post);
			mySqlCommand.Parameters.AddWithValue("@STATUS", status);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public void correctTeacherTable()
		{
			const string createTemporaryTable = "CREATE TEMPORARY TABLE teacher2 AS (SELECT * FROM teacher GROUP BY department_id, full_name)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(createTemporaryTable, mySqlConnection);
			mySqlCommand.ExecuteNonQuery();
			const string deleteTrash = "DELETE FROM teacher WHERE teacher.teacher_id NOT IN (SELECT teacher2.teacher_id FROM teacher2)";
			mySqlCommand = new MySqlCommand(deleteTrash, mySqlConnection);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}

		public int getAuditoryTypeID(string type)
		{
			const string command = "SELECT auditory_type_id FROM auditory_type WHERE auditory_type_name = @TYPE";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@TYPE", type);
			mySqlCommand.ExecuteNonQuery();
			int auditoryTypeID = Convert.ToInt32(mySqlCommand.ExecuteScalar().ToString());
			mySqlConnection.Close();
			return auditoryTypeID;	
		}
		
		public void insertAuditory(int departmentID, string name, bool not_used, int auditoryTypeID, int count, int corpsNumber)
		{
			const string command = "INSERT INTO auditory (department_id, auditory_name, not_used, type_auditory, count_of_places, corps_number) " +
				"VALUES(@ID, @AUDITORY_NAME, @NOT_USED, @TYPE_ID, @COUNT, @CORPS_NUMBER)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
			mySqlCommand.Parameters.AddWithValue("@AUDITORY_NAME", name);
			mySqlCommand.Parameters.AddWithValue("@NOT_USED", not_used);
			mySqlCommand.Parameters.AddWithValue("@TYPE_ID", auditoryTypeID);
			mySqlCommand.Parameters.AddWithValue("@COUNT", count);
			mySqlCommand.Parameters.AddWithValue("@CORPS_NUMBER", corpsNumber);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
		public void insertFullDataStudy_group(int departmentID, string code, string name, string course, int count)
		{
			const string command = "INSERT INTO study_group (department_id, study_group_code, full_name," +
			" course_number, count_of_students) VALUES(@ID, @CODE, @NAME, @COURSE, @COUNT)";
			mySqlConnection.Open();
			mySqlCommand = new MySqlCommand(command, mySqlConnection);
			mySqlCommand.Parameters.AddWithValue("@ID", departmentID);
			mySqlCommand.Parameters.AddWithValue("CODE", code);
			mySqlCommand.Parameters.AddWithValue("@NAME", name);
			mySqlCommand.Parameters.AddWithValue("@COURSE", course);
			mySqlCommand.Parameters.AddWithValue("@COUNT", count);
			mySqlCommand.ExecuteNonQuery();
			mySqlConnection.Close();
		}
		
	}
}