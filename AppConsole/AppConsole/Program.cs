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
			
			Dep_MachineBuildingTechnology mbt = new Dep_MachineBuildingTechnology("ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx", 15, 46);
			
			
			
			//auditoryTypes.SendDataToDB();
			//disciplines.SendDataToDB();
			//faculties.SendDataToDB();
			//departments.SendDataToDB();
			//teachers.SendDataToDB();
			//auditories.SendDataToDB();
			//studyGroups.SendDataToDB();
			
			//відомості
			
			//machineParts.SendDataToDB();
			mbt.SendDataToDB();
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}