using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace AppConsole
{
	public class AuditoryTypes : ExcelFile
	{		
		public AuditoryTypes(string fileName): base(fileName)
		{
			this.fileName = fileName;
		}
				
		public override void ReadFromExcelFile()
		{
			Console.WriteLine("Read");
		}
		
		public override void EvaluateData()
		{
			Console.WriteLine("EvaluateData");
		}
		
		public override void Load()
		{
			Console.WriteLine("Load");
		}
		
		
		/*public void display()
		{
			open();
			Console.WriteLine("Файл " + FileName + " відкрито для зчитування даних...");
			Console.WriteLine(rowsCount);
			close();
			Console.WriteLine("Файл " + FileName + " закрито після для зчитування даних...");
		}*/
	}
}
