using System;
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
		
		public ArrayList getDepartments()
		{
			ArrayList departments = new ArrayList();		
			mySqlConnection.Open();
			const string getDepartments = "SELECT full_name FROM department WHERE department_id NOT BETWEEN 60 AND 66";
			mySqlCommand = new MySqlCommand (getDepartments, mySqlConnection);
			using(MySqlDataReader dataReader = mySqlCommand.ExecuteReader())
			{
				while(dataReader.Read())
				{
					departments.Add(dataReader[0].ToString());
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
	}
}