using System;

using System.Collections;

namespace AppConsole
{
	class Program
	{
		public static void Main(string[] args)
		{
			
			//AuditoryTypes auditoryTypes = new AuditoryTypes("TypesOfAuditories.xlsx");	
			//Disciplines disciplines = new Disciplines("Disciplines.xlsx");
		 	//Faculties faculties = new Faculties("Faculties.xlsx");		
			//Departments departments = new Departments("Departments.xlsx");		
			//Teachers teachers = new Teachers("Teachers.xlsx");		 
		 	//Auditories auditories = new Auditories("Auditories.xls");			
			//StudyGroups studyGroups = new StudyGroups("StudyGroups.xlsx");
			

			/*
			Dep_MachineParts machineParts = new Dep_MachineParts("VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx", 15, 50);
			
			Dep_MachineBuildingTechnology mbt = new Dep_MachineBuildingTechnology("ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx", 15, 46);
			
			Dep_EconomyAndCustoms economyAndCustoms = new Dep_EconomyAndCustoms("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx", 15, 84);
			
			Dep_EconomicalTheory economicalTheory = new Dep_EconomicalTheory("ЕКОНОМІЧНОЇ ТЕОРІЇ ТА ПІДПРИЄМНИЦТВА_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 79);
			
			Dep_ElectricalMachines electricalMachines = new Dep_ElectricalMachines("Електричних_машин-Форма 44 ВІД ДОРУЧЕНЬ- 2020_кафЕМ_ден2 сем.xlsx", 15, 45);
			
			Dep_IndustrialEnergySupply industrialEnergySupply = new Dep_IndustrialEnergySupply("Електропостачання промислових підприємств_Форма 44 ЕПП - 2020д.xlsx", 15, 76);
			
			Dep_ComputerSystemsAndNetworks computerSystemsAndNetworks_sheet1 = new Dep_ComputerSystemsAndNetworks("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 93, 1);
			Dep_ComputerSystemsAndNetworks computerSystemsAndNetworks_sheet2 = new Dep_ComputerSystemsAndNetworks("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 23, 2);
			
			Dep_MarketingAndLogistics marketingAndLogistics = new Dep_MarketingAndLogistics("МАРКЕТИНГУ ТА ЛОГІСТИКИ_Відомість_денне_ІІ_нова.xls", 15, 72);
			*/
			
			Dep_InternationalEconomicRelations internationalEconomicRelations = new Dep_InternationalEconomicRelations("МІЖНАРОДНИХ ЕКОНОМІЧНИХ ВІДНОСИН_МЕВ-денне 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 57);
			
			//auditoryTypes.SendDataToDB();
			//disciplines.SendDataToDB();
			//faculties.SendDataToDB();
			//departments.SendDataToDB();
			//teachers.SendDataToDB();
			//auditories.SendDataToDB();
			//studyGroups.SendDataToDB();
			
			//відомості
			
			/*machineParts.SendDataToDB();
			mbt.SendDataToDB();
			economyAndCustoms.SendDataToDB();
			economicalTheory.SendDataToDB();
			electricalMachines.SendDataToDB();
			industrialEnergySupply.SendDataToDB();
			
			computerSystemsAndNetworks_sheet1.SendDataToDB();
			computerSystemsAndNetworks_sheet2.SendDataToDB();
			Console.WriteLine("ComputerSystemsAndNetworks Department is loaded!");
			
			marketingAndLogistics.SendDataToDB();
			*/
			
			internationalEconomicRelations.SendDataToDB();
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}