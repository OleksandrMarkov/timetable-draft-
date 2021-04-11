using System;

using System.Collections;

namespace AppConsole
{
	class Program
	{
		public static void Main(string[] args)
		{
			
			AuditoryTypes auditoryTypes = new AuditoryTypes("TypesOfAuditories.xlsx");	
			Disciplines disciplines = new Disciplines("Disciplines.xlsx");
		 	Faculties faculties = new Faculties("Faculties.xlsx");		
			Departments departments = new Departments("Departments.xlsx");		
			Teachers teachers = new Teachers("Teachers.xlsx");		 
		 	Auditories auditories = new Auditories("Auditories.xls");			
			StudyGroups studyGroups = new StudyGroups("StudyGroups.xlsx");

			Dep_MachineParts machineParts = new Dep_MachineParts("VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx", 15, 50);
			
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
			
			//auditoryTypes.SendDataToDB();
			//disciplines.SendDataToDB();
			//faculties.SendDataToDB();
			//departments.SendDataToDB();
			//teachers.SendDataToDB();
			//auditories.SendDataToDB();
			//studyGroups.SendDataToDB();
			machineParts.SendDataToDB();
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}