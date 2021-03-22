using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


/*
using Excel = Microsoft.Office.Interop.Excel;
*/

namespace TimetableApp
{
    class Program
    {
        static void Main(string[] args)
        {
            

            ArrayList excelFiles = new ArrayList();
            excelFiles.Add("Audiences.xls");
            excelFiles.Add("Departments.xlsx");
            excelFiles.Add("Disciplines.xlsx");
            excelFiles.Add("Faculties.xlsx");
            excelFiles.Add("StudyGroups.xlsx");
            excelFiles.Add("Teachers.xlsx");
            excelFiles.Add("TypesOfAudiences.xlsx");

            excelFiles.Add("VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx");
            excelFiles.Add("ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx");
            excelFiles.Add("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx");
            excelFiles.Add("ЕКОНОМІЧНОЇ ТЕОРІЇ ТА ПІДПРИЄМНИЦТВА_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            excelFiles.Add("Електричних_машин-Форма 44 ВІД ДОРУЧЕНЬ- 2020_кафЕМ_ден2 сем.xlsx");
            excelFiles.Add("Електропостачання промислових підприємств_Форма 44 ЕПП - 2020д.xlsx");
            excelFiles.Add("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx");
            excelFiles.Add("МАРКЕТИНГУ ТА ЛОГІСТИКИ_Відомість_денне_ІІ_нова.xls");
            excelFiles.Add("МІЖНАРОДНИХ ЕКОНОМІЧНИХ ВІДНОСИН_МЕВ-денне 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            excelFiles.Add("Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            excelFiles.Add("Прикладна_математика_Форма 44 ПМ денна 2019- 2020.xlsx");
            excelFiles.Add("Програмних_засобів_26-12-19_Форма 44_ ВIДОМIСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            excelFiles.Add("соціальної роботи та психології Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 Денна Соціальна робота та психологія.xlsx");
            excelFiles.Add("Технологій авіаційних двигунів ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 весна денна.xlsx");
            excelFiles.Add("Туризм_Форма 44 денна заочна 2020.xlsx");

            // создания и проверка содержимого на ошибки
            /*foreach (string excelFile in excelFiles)
             {
                 ExcelFile file = new ExcelFile(excelFile); // создать 

                //if (file.exists())

                    // if filename == "Audiences": ...
                    // if filename == "Teachers": ...

                    // некорректные значения и т.д.
                    // проверить загружены ли уже связные таблицы НА ПОТОМ!
                
            }*/

            /*ExcelFile file = new ExcelFile("TypesOfAuditories.xlsx");
            if (file.exists())
            {
                file.load();   
            }*/

            /*ExcelFile file = new ExcelFile("Disciplines.xlsx");
            if (file.exists())
            {
                file.load();
            }*/


            /*ExcelFile file = new ExcelFile("Faculties.xlsx");
            if (file.exists())
            {
                file.load();
            }*/


                /* 3. 
                 * Порядок загрузки:
                 * типы ауд
                 * дисциплины
                 * факультеты
                 * кафедры
                 * учителя
                 * аудитории
                 * группы
                 * 
                */

                Console.ReadKey();
        }
    }
}