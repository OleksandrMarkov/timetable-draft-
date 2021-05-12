using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;

namespace DataCollectionApp
{
	public partial class Window1 : Window
	{
		const string auditoryTypesText = "Типи аудиторій";
		const string disciplinesText = 	"Дисципліни";
		const string facultiesText = "Факультети";
		const string auditoriesText = "Аудиторії";
		const string departmentsText = "Кафедри";
		const string teachersText = "Викладачі";
		const string groupsText = "Учбові групи";
		
		const string machinePartsText = "Кафедра деталей машин і підйомно-транспортних механізмів";
		const string mbtText = "Кафедра технології машинобудування";
		const string economyAndCustomsText = "Кафедра економіки та митної справи";
		const string economicalTheoryText = "Кафедра економічної теорії та підприємництва";
		const string electricalMachinesText = "Кафедра електричних машин";
		const string industrialEnergySupplyText = "Кафедра електропостачання промислових підприємств";
		const string computerSystemsAndNetworksText = "Кафедра комп’ютерних систем та мереж";
		const string marketingAndLogisticsText = "Кафедра маркетингу та логістики";
		const string internationalEconomicRelationsText = "Кафедра міжнародних економічних відносин";
		const string accountingAndAuditText = "Кафедра обліку і оподаткування";
		const string appliedMathematicsText = "Кафедра прикладної математики";
		const string computerSoftwareText = "Кафедра програмних засобів";
		const string psychologyText = "Кафедра соціальної роботи";
		const string aviationEngineConstructionTechnologyText = "Кафедра технології авіаційних двигунів";
		const string tourismText =	"Кафедра туристичного, готельного та ресторанного бізнесу";
		
		AuditoryTypes auditoryTypes = new AuditoryTypes("TypesOfAuditories.xlsx");	
		Disciplines disciplines = new Disciplines("Disciplines.xlsx");
		Faculties faculties = new Faculties("Faculties.xlsx");		
		Departments departments = new Departments("Departments.xlsx");		
		Teachers teachers = new Teachers("Teachers.xlsx");		 
		Auditories auditories = new Auditories("Auditories.xls");			
		StudyGroups studyGroups = new StudyGroups("Групи 30.04.2021.xlsx");
		Dep_MachineParts machineParts = new Dep_MachineParts("VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx", 15, 50);			
		Dep_MachineBuildingTechnology mbt = new Dep_MachineBuildingTechnology("ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx", 15, 46);	
		Dep_EconomyAndCustoms economyAndCustoms_sheet1 = new Dep_EconomyAndCustoms("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx", 15, 84, 1);
		Dep_EconomyAndCustoms economyAndCustoms_sheet2 = new Dep_EconomyAndCustoms("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx", 15, 68, 2);
		Dep_EconomicalTheory economicalTheory = new Dep_EconomicalTheory("ЕКОНОМІЧНОЇ ТЕОРІЇ ТА ПІДПРИЄМНИЦТВА_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 79);
		Dep_ElectricalMachines electricalMachines = new Dep_ElectricalMachines("Електричних_машин-Форма 44 ВІД ДОРУЧЕНЬ- 2020_кафЕМ_ден2 сем.xlsx", 15, 45);
		Dep_IndustrialEnergySupply industrialEnergySupply = new Dep_IndustrialEnergySupply("Електропостачання промислових підприємств_Форма 44 ЕПП - 2020д.xlsx", 15, 76);
		Dep_ComputerSystemsAndNetworks computerSystemsAndNetworks_sheet1 = new Dep_ComputerSystemsAndNetworks("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 93, 1);
		Dep_ComputerSystemsAndNetworks computerSystemsAndNetworks_sheet2 = new Dep_ComputerSystemsAndNetworks("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 23, 2);
		Dep_MarketingAndLogistics marketingAndLogistics = new Dep_MarketingAndLogistics("МАРКЕТИНГУ ТА ЛОГІСТИКИ_Відомість_денне_ІІ_нова.xls", 15, 72);
		Dep_InternationalEconomicRelations internationalEconomicRelations = new Dep_InternationalEconomicRelations("МІЖНАРОДНИХ ЕКОНОМІЧНИХ ВІДНОСИН денне 44 2020.xlsx", 15, 57);
		Dep_AccountingAndAudit accountingAndAudit_sheet1 = new Dep_AccountingAndAudit("Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 62, 1);
		Dep_AccountingAndAudit accountingAndAudit_sheet2 = new Dep_AccountingAndAudit("Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 64, 2);
		Dep_AppliedMathematics appliedMathematics = new Dep_AppliedMathematics("Прикладна_математика_Форма 44 ПМ денна 2019- 2020.xlsx", 15, 71);
		Dep_Psychology psychology = new Dep_Psychology("Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 Денна Соціальна робота та психологія.xlsx", 15, 156);
		Dep_AviationEngineConstructionTechnology aviationEngineConstructionTechnology = new Dep_AviationEngineConstructionTechnology("Технологій авіаційних двигунів ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 весна денна.xlsx", 15, 65);
		Dep_Tourism tourism_sheet1 = new Dep_Tourism("Туризм_Форма 44 денна заочна 2020.xlsx", 15, 95, 1);
		Dep_Tourism tourism_sheet2 = new Dep_Tourism("Туризм_Форма 44 денна заочна 2020.xlsx", 15, 95, 2);	
		Dep_ComputerSoftware computerSoftware = new Dep_ComputerSoftware("Програмних_засобів_26-12-19_Форма 44_ ВIДОМIСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 225);		

		public Window1()
		{ InitializeComponent(); }
			
		void WatchButton_Click(object sender, RoutedEventArgs e)
		{
			string fileName = "";
			ListBoxItem statementsFile = (ListBoxItem)statementsListBox.SelectedItem;
			ListBoxItem commonDataFile = (ListBoxItem)commonDataListBox.SelectedItem;
			
			if(statementsFile == null && commonDataFile != null)
			{
				fileName = commonDataFile.Content.ToString();	
				switch(fileName)
				{
					case auditoryTypesText:
						auditoryTypes.openForViewing();
						break;
					case disciplinesText:
						disciplines.openForViewing();
						break;
					case facultiesText:
						faculties.openForViewing();
						break;
					case departmentsText:
						departments.openForViewing();
						break;
					case teachersText:
						teachers.openForViewing();
						break;
					case auditoriesText:
						auditories.openForViewing();
						break;
					case groupsText:
						studyGroups.openForViewing();
						break;						
				}
				fileName = "";
				commonDataListBox.SelectedItem = null;
			}
			else
			{
				if(statementsFile != null && commonDataFile == null)
				{
					fileName = statementsFile.Content.ToString();
					switch(fileName)
					{
						case machinePartsText:
							machineParts.openForViewing();
							break;			
						case mbtText:						
							mbt.openForViewing();
							break;
						case economyAndCustomsText:
							economyAndCustoms_sheet1.openForViewing();
							economyAndCustoms_sheet2.openForViewing();
							break;
						case economicalTheoryText:
							economicalTheory.openForViewing();
							break;
						case electricalMachinesText:
							electricalMachines.openForViewing();
							break;
						case industrialEnergySupplyText:
							industrialEnergySupply.openForViewing();
							break;
						case computerSystemsAndNetworksText:
							computerSystemsAndNetworks_sheet1.openForViewing();
							computerSystemsAndNetworks_sheet2.openForViewing();
							break;
						case marketingAndLogisticsText:
							marketingAndLogistics.openForViewing();
							break;
						case internationalEconomicRelationsText:
							internationalEconomicRelations.openForViewing();
							break;
						case accountingAndAuditText:
							accountingAndAudit_sheet1.openForViewing();
							accountingAndAudit_sheet2.openForViewing();
							break;
						case appliedMathematicsText:
							appliedMathematics.openForViewing();
							break;
						case computerSoftwareText:
							computerSoftware.openForViewing();
							break;
						case psychologyText:
							psychology.openForViewing();
							break;
						case aviationEngineConstructionTechnologyText:
							aviationEngineConstructionTechnology.openForViewing();
							break;
						case tourismText:
							tourism_sheet1.openForViewing();
							tourism_sheet2.openForViewing();
							break;							
					}							
					fileName = "";
					statementsListBox.SelectedItem = null;
				}
				else 
				{
					if(statementsFile == null && commonDataFile == null)
					{ MessageBox.Show("Не обрано файл для перегляду!");	}
				}
			}		
		}

		void LoadButton_Click(object sender, RoutedEventArgs e)
		{
			string fileName = "";
			ListBoxItem statementsFile = (ListBoxItem)statementsListBox.SelectedItem;
			ListBoxItem commonDataFile = (ListBoxItem)commonDataListBox.SelectedItem;
			
			if(statementsFile == null && commonDataFile != null)
			{
				fileName = commonDataFile.Content.ToString();
				switch(fileName)
				{
					case auditoryTypesText:
						auditoryTypes.SendDataToDB();
						break;
					case disciplinesText:
						disciplines.SendDataToDB();
						break;
					case facultiesText:
						faculties.SendDataToDB();
						break;
					case departmentsText:
						departments.SendDataToDB();
						MessageBox.Show("Дані про кафедри завантажено до бази даних!");
						lbi_departments.IsEnabled = false;
						lbi_departments.FontWeight = FontWeights.Bold;
						break;
					case teachersText:
						teachers.SendDataToDB();
						MessageBox.Show("Дані про викладачів завантажено до бази даних!");
						lbi_teachers.IsEnabled = false;
						lbi_teachers.FontWeight = FontWeights.Bold;
						break;
					case auditoriesText:
						auditories.SendDataToDB();
						MessageBox.Show("Дані про аудиторії завантажено до бази даних!");
						lbi_auditories.IsEnabled = false;
						lbi_auditories.FontWeight = FontWeights.Bold;
						break;
					case groupsText:
						studyGroups.SendDataToDB();
						MessageBox.Show("Дані про учбові групи завантажено до бази даних!");
						lbi_groups.IsEnabled = false;
						lbi_groups.FontWeight = FontWeights.Bold;
						break;						
				}
				fileName = "";
				commonDataListBox.SelectedItem = null;
			}
			else
			{
				if(statementsFile != null && commonDataFile == null)
				{
					fileName = statementsFile.Content.ToString();
					switch(fileName)
					{
						case machinePartsText:						
							machineParts.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри деталей машин " +
							"і підйомно-транспортних механізмів завантажено до бази даних!");
							lbi_mp.IsEnabled = false;
							lbi_mp.FontWeight = FontWeights.Bold;
							break;		
						case mbtText:						
							mbt.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри технології машинобудування " +
							"завантажено до бази даних!");
							lbi_mbt.IsEnabled = false;
							lbi_mbt.FontWeight = FontWeights.Bold;
							break;
						case economyAndCustomsText:
							economyAndCustoms_sheet1.SendDataToDB();
							economyAndCustoms_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри економіки та митної справи " +
							"завантажено до бази даних!");
							lbi_eac.IsEnabled = false;
							lbi_eac.FontWeight = FontWeights.Bold;
							break;
						case economicalTheoryText:
							economicalTheory.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри економічної теорії та підприємництва " +
							"завантажено до бази даних!");
							lbi_et.IsEnabled = false;
							lbi_et.FontWeight = FontWeights.Bold;
							break;
						case electricalMachinesText:
							electricalMachines.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри електричних машин " +
							"завантажено до бази даних!");
							lbi_em.IsEnabled = false;
							lbi_em.FontWeight = FontWeights.Bold;
							break;
						case industrialEnergySupplyText:
							industrialEnergySupply.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри електропостачання промислових підприємств " +
							"завантажено до бази даних!");
							lbi_ies.IsEnabled = false;
							lbi_ies.FontWeight = FontWeights.Bold;		
							break;
						case computerSystemsAndNetworksText:
							computerSystemsAndNetworks_sheet1.SendDataToDB();
							computerSystemsAndNetworks_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри комп’ютерних систем та мереж " +
							"завантажено до бази даних!");
							lbi_csan.IsEnabled = false;
							lbi_csan.FontWeight = FontWeights.Bold;
							break;
						case marketingAndLogisticsText:
							marketingAndLogistics.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри маркетингу та логістики " +
							"завантажено до бази даних!");
							lbi_mal.IsEnabled = false;
							lbi_mal.FontWeight = FontWeights.Bold;							
							break;
						case internationalEconomicRelationsText:
							internationalEconomicRelations.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри міжнародних економічних відносин " +
							"завантажено до бази даних!");
							lbi_ier.IsEnabled = false;
							lbi_ier.FontWeight = FontWeights.Bold;
							break;
						case accountingAndAuditText:
							accountingAndAudit_sheet1.SendDataToDB();
							accountingAndAudit_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри обліку і оподаткування " +
							"завантажено до бази даних!");
							lbi_aaa.IsEnabled = false;
							lbi_aaa.FontWeight = FontWeights.Bold;
							break;
						case appliedMathematicsText:
							appliedMathematics.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри прикладної математики " +
							"завантажено до бази даних!");
							lbi_am.IsEnabled = false;
							lbi_am.FontWeight = FontWeights.Bold;
							break;
						case computerSoftwareText:
							computerSoftware.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри програмних засобів " +
							"завантажено до бази даних!");
							lbi_cs.IsEnabled = false;
							lbi_cs.FontWeight = FontWeights.Bold;
							break;
						case psychologyText:
							psychology.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри соціальної роботи " +
							"завантажено до бази даних!");
							lbi_ps.IsEnabled = false;
							lbi_ps.FontWeight = FontWeights.Bold;
							break;
						case aviationEngineConstructionTechnologyText:
							aviationEngineConstructionTechnology.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри технології авіаційних двигунів " +
							"завантажено до бази даних!");
							lbi_aect.IsEnabled = false;
							lbi_aect.FontWeight = FontWeights.Bold;
							break;
						case tourismText:
							tourism_sheet1.SendDataToDB();
							tourism_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри туристичного," +
"							готельного та ресторанного бізнесу " +
							"завантажено до бази даних!");
							lbi_t.IsEnabled = false;
							lbi_t.FontWeight = FontWeights.Bold;
							break;							
					}							
					fileName = "";
					statementsListBox.SelectedItem = null;
				}
				else 
				{
					if(statementsFile == null && commonDataFile == null)
					{
						MessageBox.Show("Не обрано дані для завантаження!");	
					}
				}
			}
		}		
		
		void WatchBugsReport(object sender, RoutedEventArgs e)
		{
			const string path = @"E:\BACHELORS WORK\TIMETABLE\DataCollectionApp\BugsReport.txt";
			Process.Start(path);
			
		}
		
		void lbi_auditoryTypes_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = auditoryTypesText;
			selectedFileName.Text = auditoryTypes.FileName;
			lastWriteTime.Text = auditoryTypes.LastWriteTime;
		}
		void lbi_disciplines_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = disciplinesText;
			selectedFileName.Text = disciplines.FileName;
			lastWriteTime.Text = disciplines.LastWriteTime;			
		}		
		void lbi_faculties_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = facultiesText;
			selectedFileName.Text = faculties.FileName;
			lastWriteTime.Text = faculties.LastWriteTime;			
		}
		void lbi_departments_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = departmentsText;
			selectedFileName.Text = departments.FileName;
			lastWriteTime.Text = departments.LastWriteTime;			
		}		
		void lbi_teachers_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = teachersText;
			selectedFileName.Text = teachers.FileName;
			lastWriteTime.Text = teachers.LastWriteTime;			
		}
		void lbi_auditories_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = auditoriesText;
			selectedFileName.Text = auditories.FileName;
			lastWriteTime.Text = auditories.LastWriteTime;			
		}
		void lbi_groups_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = groupsText;
			selectedFileName.Text = studyGroups.FileName;
			lastWriteTime.Text = studyGroups.LastWriteTime;			
		}		
		
		void lbi_mp_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = machinePartsText;
			selectedFileName.Text = machineParts.FileName;
			lastWriteTime.Text = machineParts.LastWriteTime;			
		}		
		void lbi_mbt_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = mbtText;
			selectedFileName.Text = mbt.FileName;
			lastWriteTime.Text = mbt.LastWriteTime;			
		}		
		void lbi_eac_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = economyAndCustomsText;
			selectedFileName.Text = economyAndCustoms_sheet1.FileName;
			lastWriteTime.Text = String.Compare(economyAndCustoms_sheet1.LastWriteTime, economyAndCustoms_sheet2.LastWriteTime)  >= 0 ?
				economyAndCustoms_sheet1.LastWriteTime : economyAndCustoms_sheet2.LastWriteTime;
		}		
		void lbi_et_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = economicalTheoryText;
			selectedFileName.Text = economicalTheory.FileName;
			lastWriteTime.Text = economicalTheory.LastWriteTime;
		}	
		void lbi_em_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = electricalMachinesText;
			selectedFileName.Text = electricalMachines.FileName;
			lastWriteTime.Text = electricalMachines.LastWriteTime;			
		}		
		void lbi_ies_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = industrialEnergySupplyText;
			selectedFileName.Text = industrialEnergySupply.FileName;
			lastWriteTime.Text = industrialEnergySupply.LastWriteTime;			
		}
		void lbi_csan_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = computerSystemsAndNetworksText;
			selectedFileName.Text = computerSystemsAndNetworks_sheet1.FileName;
			lastWriteTime.Text = String.Compare(computerSystemsAndNetworks_sheet1.LastWriteTime, computerSystemsAndNetworks_sheet2.LastWriteTime)  >= 0 ?
				computerSystemsAndNetworks_sheet1.LastWriteTime : computerSystemsAndNetworks_sheet2.LastWriteTime;			
		}
		void lbi_mal_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = marketingAndLogisticsText;
			selectedFileName.Text = marketingAndLogistics.FileName;
			lastWriteTime.Text = marketingAndLogistics.LastWriteTime;			
		}		
		void lbi_ier_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = internationalEconomicRelationsText;
			selectedFileName.Text = internationalEconomicRelations.FileName;
			lastWriteTime.Text = internationalEconomicRelations.LastWriteTime;			
		}
		void lbi_aaa_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = accountingAndAuditText;
			selectedFileName.Text = accountingAndAudit_sheet1.FileName;
			lastWriteTime.Text = String.Compare(accountingAndAudit_sheet1.LastWriteTime, accountingAndAudit_sheet2.LastWriteTime)  >= 0 ?
				accountingAndAudit_sheet1.LastWriteTime : accountingAndAudit_sheet2.LastWriteTime;			
		}
		void lbi_am_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = appliedMathematicsText;
			selectedFileName.Text = appliedMathematics.FileName;
			lastWriteTime.Text = appliedMathematics.LastWriteTime;			
		}
		void lbi_cs_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = computerSoftwareText;
			selectedFileName.Text = computerSoftware.FileName;
			lastWriteTime.Text = computerSoftware.LastWriteTime;			
		}
		void lbi_ps_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = psychologyText;
			selectedFileName.Text = psychology.FileName;
			lastWriteTime.Text = psychology.LastWriteTime;			
		}
		void lbi_aect_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = aviationEngineConstructionTechnologyText;
			selectedFileName.Text = aviationEngineConstructionTechnology.FileName;
			lastWriteTime.Text = aviationEngineConstructionTechnology.LastWriteTime;			
		}
		void lbi_t_selected(object sender, RoutedEventArgs e)
		{
			selectedFile.Text = tourismText;
			selectedFileName.Text = tourism_sheet1.FileName;
			lastWriteTime.Text = String.Compare(tourism_sheet1.LastWriteTime, tourism_sheet2.LastWriteTime)  >= 0 ?
				tourism_sheet1.LastWriteTime : tourism_sheet2.LastWriteTime;			
		}
	}
}