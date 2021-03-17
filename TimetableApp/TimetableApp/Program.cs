using System;
using System.Collections.Generic;
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
        enum GeneralInformation
        {
            Audiences,
            Departments,
            Disciplines,
            Faculties,
            StudyGroups,
            Teachers,
            TypesOfAudiences
        }

        enum DepartmentsStatements
        {
            MachinePartsAndWindingMechanisms,
            MachineBuildingTechnology,
            EconomyAndCustoms,
            EconomicalTheoryAndEntrepreneurship,
            ElectricalMachines,
            IndustrialEnergySupply,
            ComputerSystemsAndNetworks,
            MarketingAndLogistics,
            InternationalEconomicRelations,
            AccountingAndAudit,
            AppliedMathematics,
            ComputerSoftware,
            Psychology,
            AviationEngineConstructionTechnology,
            InternationalTourism
        }

        static void Main(string[] args)
        {
            Dictionary <GeneralInformation, string> FilesWithGeneralInformation = new Dictionary<GeneralInformation, string>();
            FilesWithGeneralInformation.Add(GeneralInformation.Audiences, "Audiences.xls");
            FilesWithGeneralInformation.Add(GeneralInformation.Departments, "Departments.xlsx");
            FilesWithGeneralInformation.Add(GeneralInformation.Disciplines, "Disciplines.xlsx");
            FilesWithGeneralInformation.Add(GeneralInformation.Faculties, "Faculties.xlsx");
            FilesWithGeneralInformation.Add(GeneralInformation.StudyGroups, "StudyGroups.xlsx");
            FilesWithGeneralInformation.Add(GeneralInformation.Teachers, "Teachers.xlsx");
            FilesWithGeneralInformation.Add(GeneralInformation.TypesOfAudiences, "TypesOfAudiences.xlsx");

            Dictionary<DepartmentsStatements, string> FilesWithDepartmentsStatements = new Dictionary<DepartmentsStatements, string>();
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.MachinePartsAndWindingMechanisms, "VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.MachineBuildingTechnology, "ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.EconomyAndCustoms, "Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.EconomicalTheoryAndEntrepreneurship, "ЕКОНОМІЧНОЇ ТЕОРІЇ ТА ПІДПРИЄМНИЦТВА_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.ElectricalMachines, "Електричних_машин-Форма 44 ВІД ДОРУЧЕНЬ- 2020_кафЕМ_ден2 сем.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.IndustrialEnergySupply, "Електропостачання промислових підприємств_Форма 44 ЕПП - 2020д.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.ComputerSystemsAndNetworks, "КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.MarketingAndLogistics, "МАРКЕТИНГУ ТА ЛОГІСТИКИ_Відомість_денне_ІІ_нова.xls");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.InternationalEconomicRelations, "МІЖНАРОДНИХ ЕКОНОМІЧНИХ ВІДНОСИН_МЕВ-денне 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.AccountingAndAudit, "Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.AppliedMathematics, "Прикладна_математика_Форма 44 ПМ денна 2019- 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.ComputerSoftware, "Програмних_засобів_26-12-19_Форма 44_ ВIДОМIСТЬ ДОРУЧЕНЬ - 2020.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.Psychology, "соціальної роботи та психології Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 Денна Соціальна робота та психологія.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.AviationEngineConstructionTechnology, "Технологій авіаційних двигунів ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 весна денна.xlsx");
            FilesWithDepartmentsStatements.Add(DepartmentsStatements.InternationalTourism, "Туризм_Форма 44 денна заочна 2020.xlsx");

            // создания и проверка содержимого на ошибки
            foreach (KeyValuePair<GeneralInformation, string> keyValue in FilesWithGeneralInformation)
            {
                ExcelFile file = new ExcelFile(keyValue.Value); // создать 
                
                
                // проверить на дубликаты, пустые значения, некорректные значения и т.д.

                // проверить загружены ли уже связные таблицы
            }

            /* Порядок загрузки:
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