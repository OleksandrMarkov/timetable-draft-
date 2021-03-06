﻿using System;
using System.Collections;
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
		
		const string machinePartsText = "Деталей машин і підйомно-транспортних механізмів";
		const string mbtText = "Технології машинобудування";
		const string economyAndCustomsText = "Економіки та митної справи";
		const string economicalTheoryText = "Економічної теорії та підприємництва";
		const string electricalMachinesText = "Електричних машин";
		const string industrialEnergySupplyText = "Електропостачання промислових підприємств";
		const string computerSystemsAndNetworksText = "Комп'ютерних систем та мереж";
		const string marketingAndLogisticsText = "Маркетингу та логістики";
		const string internationalEconomicRelationsText = "Міжнародних економічних відносин";
		const string accountingAndAuditText = "Обліку і оподаткування";
		const string appliedMathematicsText = "Прикладної математики";
		const string computerSoftwareText = "Програмних засобів";
		const string psychologyText = "Соціальної роботи та психології";
		const string aviationEngineConstructionTechnologyText = "Технології авіаційних двигунів";
		const string tourismText =	"Туристичного, готельного та ресторанного бізнесу";
		
		AuditoryTypes auditoryTypes = new AuditoryTypes("TypesOfAuditories.xlsx");	
		Disciplines disciplines = new Disciplines("Disciplines.xlsx");
		Faculties faculties = new Faculties("Faculties.xlsx");		
		Departments departments = new Departments("Departments.xlsx");		
		Teachers teachers = new Teachers("Teachers.xlsx");		 
		Auditories auditories = new Auditories("Auditories_недостающие_корпуса_из_путеводителя_ЗНТУ_мои_предположения.xls");			
		StudyGroups studyGroups = new StudyGroups("Групи 30.04.2021.xlsx");

		Statements appliedMathematics = new Statements("Прикладна_математика_Форма 44 ПМ денна 2019- 2020.xlsx", 15, 1, "ПМ");
		Statements internationalEconomicRelations = new Statements("МІЖНАРОДНИХ ЕКОНОМІЧНИХ ВІДНОСИН денне 44 2020.xlsx", 15, 1, "МЕВ");
		Statements machineParts = new Statements("VIDOMOST_DORUChEN_2 сем_ДВ_ДМ і ПТМ.xlsx", 15, 1, "ДМіПТМ");
		Statements mbt = new Statements("ВІДОМІСТЬ ДОРУЧЕНЬ ТМБ денне весна - 2020.xlsx", 15, 1, "ТМБ");
		Statements economyAndCustoms_sheet1 = new Statements("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx", 15, 1, "ЕтаМС");
		Statements economyAndCustoms_sheet2 = new Statements("Економіки та митної справи_Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020_ЕМС.xlsx", 15, 2, "ЕтаМС");		
		Statements economicalTheory = new Statements("ЕКОНОМІЧНОЇ ТЕОРІЇ ТА ПІДПРИЄМНИЦТВА_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 1, "ЕТтаП");
		Statements electricalMachines = new Statements("Електричних_машин-Форма 44 ВІД ДОРУЧЕНЬ- 2020_кафЕМ_ден2 сем.xlsx", 15, 1, "ЕМ");
		Statements industrialEnergySupply = new Statements("Електропостачання промислових підприємств_Форма 44 ЕПП - 2020д.xlsx", 15, 1, "ЕПП");
		Statements computerSystemsAndNetworks_sheet1 = new Statements("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 1, "КСтаМ");
		Statements computerSystemsAndNetworks_sheet2 = new Statements("КОМП_ЮТЕРНІ СИСТЕМИ ТА МЕРЕЖІ_ВІДОМІСТЬ ДОРУЧЕНЬ_19_20.xlsx", 15, 2, "КСтаМ");		
		Statements marketingAndLogistics = new Statements("МАРКЕТИНГУ ТА ЛОГІСТИКИ_Відомість_денне_ІІ_нова.xls", 15, 1, "Марк.та Лог.");
		Statements accountingAndAudit_sheet1 = new Statements("Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 1, "ОіО");
		Statements accountingAndAudit_sheet2 = new Statements("Облік і оподатківання_ВІДОМІСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 2, "ОіО");
		Statements psychology = new Statements("Форма 44 ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 Денна Соціальна робота та психологія.xlsx", 15, 1, "СоцРтаП");
		Statements aviationEngineConstructionTechnology = new Statements("Технологій авіаційних двигунів ВІДОМІСТЬ ДОРУЧЕНЬ - 2020 весна денна.xlsx", 15, 1, "ТАД");		
		Statements computerSoftware = new Statements("Програмних_засобів_26-12-19_Форма 44_ ВIДОМIСТЬ ДОРУЧЕНЬ - 2020.xlsx", 15, 1, "ПЗ");
 		Statements tourism_sheet1 = new Statements("Туризм_Форма 44 денна заочна 2020.xlsx", 15, 1, "ТГтаРБ");
		Statements tourism_sheet2 = new Statements("Туризм_Форма 44 денна заочна 2020.xlsx", 15, 2, "ТГтаРБ");
						
		public Window1()
		{ 
			InitializeComponent();
			
			DbOperations dbo = new DbOperations();
			Dictionary<int, string> departments = dbo.getDepartments();
			foreach (var d in departments)
			{
				ListBoxItem lbi = new ListBoxItem();
				lbi.Content = d.Value;
				lbi.Uid = d.Key.ToString();
				statementsListBox.Items.Add(lbi);
			}
		}
			
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
							//MessageBox.Show(machineParts.lastRow.ToString());
							break;			
						case mbtText:						
							mbt.openForViewing();
							//MessageBox.Show(mbt.lastRow.ToString());
							break;
						case economyAndCustomsText:
							//MessageBox.Show(economyAndCustoms_sheet1.lastRow.ToString());
							//MessageBox.Show(economyAndCustoms_sheet2.lastRow.ToString());
							economyAndCustoms_sheet1.openForViewing();
							economyAndCustoms_sheet2.openForViewing();
							break;
						case economicalTheoryText:
							//MessageBox.Show(economicalTheory.lastRow.ToString());
							economicalTheory.openForViewing();
							break;
						case electricalMachinesText:
							//MessageBox.Show(electricalMachines.lastRow.ToString());
							electricalMachines.openForViewing();
							break;
						case industrialEnergySupplyText:
							//MessageBox.Show(industrialEnergySupply.lastRow.ToString());
							industrialEnergySupply.openForViewing();
							break;
						case computerSystemsAndNetworksText:
							//MessageBox.Show(computerSystemsAndNetworks_sheet1.lastRow.ToString());
							//MessageBox.Show(computerSystemsAndNetworks_sheet2.lastRow.ToString());
							computerSystemsAndNetworks_sheet1.openForViewing();
							computerSystemsAndNetworks_sheet2.openForViewing();
							break;
						case marketingAndLogisticsText:
							//MessageBox.Show(marketingAndLogistics.lastRow.ToString());
							marketingAndLogistics.openForViewing();
							break;
						case internationalEconomicRelationsText:
							//MessageBox.Show(internationalEconomicRelations.lastRow.ToString());
							internationalEconomicRelations.openForViewing();
							break;
						case accountingAndAuditText:
							//MessageBox.Show(accountingAndAudit_sheet1.lastRow.ToString());
							//MessageBox.Show(accountingAndAudit_sheet2.lastRow.ToString());
							accountingAndAudit_sheet1.openForViewing();
							accountingAndAudit_sheet2.openForViewing();
							break;
						case appliedMathematicsText:
							//MessageBox.Show(appliedMathematics.lastRow.ToString());
							appliedMathematics.openForViewing();				
							break;
						case computerSoftwareText:
							//MessageBox.Show(computerSoftware.lastRow.ToString());
							computerSoftware.openForViewing();
							break;
						case psychologyText:
							//MessageBox.Show(psychology.lastRow.ToString());
							psychology.openForViewing();
							break;
						case aviationEngineConstructionTechnologyText:
							//MessageBox.Show(aviationEngineConstructionTechnology.lastRow.ToString());							
							aviationEngineConstructionTechnology.openForViewing();
							break;
						case tourismText:
							//MessageBox.Show(tourism_sheet1.lastRow.ToString());
							//MessageBox.Show(tourism_sheet2.lastRow.ToString());							
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
						break;
					case teachersText:
						teachers.SendDataToDB();
						break;
					case auditoriesText:
						auditories.SendDataToDB();
						break;
					case groupsText:
						studyGroups.SendDataToDB();
						break;						
				}
				//MessageBox.Show("Дані про " + fileName.ToLower() + " завантажено до бази даних!");
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
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;		
						case mbtText:						
							mbt.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case economyAndCustomsText:
							economyAndCustoms_sheet1.SendDataToDB();
							economyAndCustoms_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case economicalTheoryText:
							economicalTheory.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case electricalMachinesText:
							electricalMachines.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case industrialEnergySupplyText:
							industrialEnergySupply.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case computerSystemsAndNetworksText:
							computerSystemsAndNetworks_sheet1.SendDataToDB();
							computerSystemsAndNetworks_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case marketingAndLogisticsText:
							marketingAndLogistics.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");							
							break;
						case internationalEconomicRelationsText:
							internationalEconomicRelations.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case accountingAndAuditText:
							accountingAndAudit_sheet1.SendDataToDB();
							accountingAndAudit_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");						
							break;
						case appliedMathematicsText:
							appliedMathematics.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");		
							break;
						case computerSoftwareText:
							computerSoftware.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");							
							break;
						case psychologyText:
							psychology.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case aviationEngineConstructionTechnologyText:
							aviationEngineConstructionTechnology.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						case tourismText:
							tourism_sheet1.SendDataToDB();
							tourism_sheet2.SendDataToDB();
							MessageBox.Show("Відомості доручень кафедри " + fileName +
					" завантажено до бази даних!");
							break;
						default:
							MessageBox.Show("Відомості доручень цієї кафедри ще не завантажили...");
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
		
		void statementsListBox_Changed(object sender, SelectionChangedEventArgs e)
		{			
			try
			{
				selectedFile.Text = ((ListBoxItem)statementsListBox.SelectedItem).Content.ToString();	
			}
			catch (Exception ex)
			{
				selectedFile.Text = null;
			}
			
			switch(selectedFile.Text)
			{
				case machinePartsText:
					selectedFileName.Text = machineParts.FileName;
					lastWriteTime.Text = machineParts.LastWriteTime;
					break;			
				case mbtText:						
					selectedFileName.Text = mbt.FileName;
					lastWriteTime.Text = mbt.LastWriteTime;
					break;
				case economyAndCustomsText:
					selectedFileName.Text =	economyAndCustoms_sheet1.FileName;
					lastWriteTime.Text = economyAndCustoms_sheet1.LastWriteTime;
					break;
				case economicalTheoryText:
					selectedFileName.Text = economicalTheory.FileName;
					lastWriteTime.Text = economicalTheory.LastWriteTime;
					break;
				case electricalMachinesText:
					selectedFileName.Text = electricalMachines.FileName;
					lastWriteTime.Text = electricalMachines.LastWriteTime;
					break;
				case industrialEnergySupplyText:
					selectedFileName.Text = industrialEnergySupply.FileName;
					lastWriteTime.Text = industrialEnergySupply.LastWriteTime;
					break;
				case computerSystemsAndNetworksText:
					selectedFileName.Text = computerSystemsAndNetworks_sheet1.FileName;
					lastWriteTime.Text = computerSystemsAndNetworks_sheet1.LastWriteTime;
					break;
				case marketingAndLogisticsText:
					selectedFileName.Text = marketingAndLogistics.FileName;
					lastWriteTime.Text = marketingAndLogistics.LastWriteTime;
					break;
				case internationalEconomicRelationsText:
					selectedFileName.Text = internationalEconomicRelations.FileName;
					lastWriteTime.Text = internationalEconomicRelations.LastWriteTime;
					break;
				case accountingAndAuditText:
					selectedFileName.Text = accountingAndAudit_sheet1.FileName;
					lastWriteTime.Text = accountingAndAudit_sheet1.LastWriteTime;
					break;
				case appliedMathematicsText:
					selectedFileName.Text = appliedMathematics.FileName;
					lastWriteTime.Text = appliedMathematics.LastWriteTime;
					break;
				case computerSoftwareText:
					selectedFileName.Text = computerSoftware.FileName;
					lastWriteTime.Text = computerSoftware.LastWriteTime;
					break;
				case psychologyText:
					selectedFileName.Text = psychology.FileName;
					lastWriteTime.Text = psychology.LastWriteTime;
					break;
				case aviationEngineConstructionTechnologyText:
					selectedFileName.Text = aviationEngineConstructionTechnology.FileName;
					lastWriteTime.Text = aviationEngineConstructionTechnology.LastWriteTime;
					break;
				case tourismText:
					selectedFileName.Text = tourism_sheet1.FileName;
					lastWriteTime.Text = tourism_sheet1.LastWriteTime;
					break;
				default:
					selectedFileName.Text = "Відомості доручень цієї кафедри ще не завантажили...";
					lastWriteTime.Text = "...";
					break;
			}
		}
		
		void commonDataListBox_Changed(object sender, SelectionChangedEventArgs e)
		{
			try
			{
				selectedFile.Text = ((ListBoxItem)commonDataListBox.SelectedItem).Content.ToString();	
			}
			catch (Exception ex)
			{
				selectedFile.Text = null;
			}
			switch(selectedFile.Text)
			{
					case auditoryTypesText:
						selectedFileName.Text = auditoryTypes.FileName;
						lastWriteTime.Text = auditoryTypes.LastWriteTime;
						break;
					case disciplinesText:
						selectedFileName.Text = disciplines.FileName;
						lastWriteTime.Text = disciplines.LastWriteTime;
						break;
					case facultiesText:
						selectedFileName.Text = faculties.FileName;
						lastWriteTime.Text = faculties.LastWriteTime;
						break;
					case departmentsText:
						selectedFileName.Text = departments.FileName;
						lastWriteTime.Text = departments.LastWriteTime;
						break;
					case teachersText:
						selectedFileName.Text = teachers.FileName;
						lastWriteTime.Text = teachers.LastWriteTime;
						break;
					case auditoriesText:
						selectedFileName.Text = auditories.FileName;
						lastWriteTime.Text = auditories.LastWriteTime;
						break;
					case groupsText:
						selectedFileName.Text = studyGroups.FileName;
						lastWriteTime.Text = studyGroups.LastWriteTime;
						break;				
			}			
		}
	}
}