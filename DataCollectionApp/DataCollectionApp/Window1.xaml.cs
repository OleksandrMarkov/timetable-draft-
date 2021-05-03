using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace DataCollectionApp
{
	public partial class Window1 : Window
	{
		public Window1()
		{
			InitializeComponent();
		}

		void WatchButton_Click(object sender, RoutedEventArgs e)
		{
			// доступ к контенту ListBox
		/*	if(statementsListBox.SelectedItem == null)
			{
				MessageBox.Show("select something!");
			}
			else
			{
				ListBoxItem lbi = (ListBoxItem)statementsListBox.SelectedItem;
				MessageBox.Show("watch " + lbi.Content);
				statementsListBox.SelectedItem = null;
			}*/
			
		}

		void LoadButton_Click(object sender, RoutedEventArgs e)
		{
			MessageBox.Show("load");
		}		
		
		void DeleteButton_Click(object sender, RoutedEventArgs e)
		{
			MessageBox.Show("delete");
		}
	}
}