using System;

using System.Collections;

namespace AppConsole
{
	class Program
	{
		public static void Main(string[] args)
		{
			
			AuditoryTypes auditoryTypes = new AuditoryTypes("TypesOfAuditories.xlsx");
		    //file.SendDataToDB();
		
			Disciplines disciplines = new Disciplines("Disciplines.xlsx");
		 	//file2.SendDataToDB();
		 
		 	Faculties faculties = new Faculties("Faculties.xlsx");
			//file3.SendDataToDB();
			
			Departments departments = new Departments("Departments.xlsx");
			//file4.SendDataToDB();
		
			Teachers teachers = new Teachers("Teachers.xlsx");
			//file5.SendDataToDB();
		 
		 	Auditories auditories = new Auditories("Auditories.xls");
			//file6.SendDataToDB();	
			
			StudyGroups studyGroups = new StudyGroups("StudyGroups.xlsx");
			//file7.SendDataToDB();
					
			/*ArrayList excelFiles = new ArrayList();
			excelFiles.Add(auditoryTypes);
			excelFiles.Add(disciplines);
			excelFiles.Add(faculties);
			excelFiles.Add(departments);
			excelFiles.Add(teachers);
			excelFiles.Add(auditories);
			excelFiles.Add(studyGroups);
			
			foreach(ExcelFile excelFile in excelFiles)
			{
				excelFile.SendDataToDB();
			}*/
			
			teachers.SendDataToDB();
			
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}