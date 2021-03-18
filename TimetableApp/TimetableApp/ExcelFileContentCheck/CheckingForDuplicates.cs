using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace TimetableApp
{
    class CheckingForDuplicates : ExcelFile
    {
        public CheckingForDuplicates(string file) : base(file)
        {
            
        }

        public void defineDataCollections()
        {
            ArrayList recordsList = new ArrayList();
            Dictionary<int, string> duplicates = new Dictionary<int, string>();

            if (FileName == "TypesOfAudiences.xlsx")
            {
                columnForReading = 'A';
            }
            
        }

        public void getDataFromFile()
        {

        }

                  /*      try
                {
                    open();
                    for (int row = 1, column = getColumnNumber(columnForReading); row <= rowsCount; row++)
                    {
                        //Console.WriteLine(row);
                        cellContent = ((range.Cells[row, column] as Excel.Range).Value2).ToString();
                        //Console.WriteLine("cellContent: " + cellContent);
                        if (recordsList.Contains(cellContent) == true)
                        {
                            duplicates.Add(row, cellContent);
                        }
                        else
                        {
                            recordsList.Add(cellContent);
                        }
                    }
                    close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Помилка при зчитуванні даних з файлу " + FileName + " " + ex.Message);
                }*/
























        /*public override bool isTrash()
        {
            return true;
        }*/
    }
}
