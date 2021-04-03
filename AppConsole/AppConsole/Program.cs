using System;

namespace AppConsole
{
	class Program
	{
		public static void Main(string[] args)
		{
			//Console.WriteLine("Hello World!");
			
			//ExcelFile excelFile = new ExcelFile("TypesOfAuditories.xlsx");
			//Console.WriteLine(excelFile.FullPathToFile);
			
			AuditoryTypes file = new AuditoryTypes("TypesOfAuditories.xlsx");
			//Console.WriteLine(file.FileName + "\n" + file.FullPathToFile);
			
			file.SendDataToDB();
			
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}