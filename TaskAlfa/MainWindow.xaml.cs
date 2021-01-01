using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TaskAlfa.Models;

namespace TaskAlfa
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public DataFile Value { get; set; }
        public DataFile Value1 { get; set; }

        

        private DataModel dm1 { get; set; }
        private DataModel dm2 { get; set; }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "File Excel| *.XLSX; *.XLS; *.CSV";
            if (fileDialog.ShowDialog().Value)
            {
                string path = fileDialog.FileName;
                var reader = ReaderExcelData.Factory.CreatReader(path);
                dm1 = reader.Read(path);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "File Excel| *.XLSX; *.XLS; *.CSV";
            if (fileDialog.ShowDialog().Value)
            {
                string path = fileDialog.FileName;
                var reader = ReaderExcelData.Factory.CreatReader(path);
                dm2 = reader.Read(path);
            }
        }
        
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            List<string> result = new List<string>();
            foreach(var account in dm1.accounts)
            {
                foreach(var currency in account.Value.currency)
                {
                    double diff = dm2.getValue(account.Key, currency.Key) - currency.Value;
                    if(diff !=0)
                    result.Add($"{account.Key}-{currency.Key} diff: {diff}");
                }
            }
            CsvData csD = new CsvData();
            csD.Save(@"C:\Users\Andrew\Desktop\TaskChallenge\answer.XLSX", result);
            MessageBox.Show("Done");
        }
    }
}
