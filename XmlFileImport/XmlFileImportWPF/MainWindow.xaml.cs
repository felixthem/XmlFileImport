using System;
using System.IO;
using System.Windows;
using XmlFileImportBase;
using XmlFileImportBase.Helpers;
using XmlFileImportBase.Models;

namespace XmlFileImportWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private channel ch = null;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void readData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string name = "data.xml";
                ch = XmlHelper.ParseXML(name);

                MessageBox.Show($@"Данные успешно считаны из файла: {name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }

        private void writeTxt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ch == null)
                {
                    throw new Exception($@"Перед записью необходимо считать данные!");
                }
                string name = "txtFile.txt";
                TaskHelper.RunTaskWriteTxt(ch, name);

                MessageBox.Show($@"Данные успешно записаны в файл: {name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }

        private void writeWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ch == null)
                {
                    throw new Exception($@"Перед записью необходимо считать данные!");
                }
                string name = "wordFile.doc";
                TaskHelper.RunTaskWriteTxt(ch, name);

                MessageBox.Show($@"Данные успешно записаны в файл: {name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }

        private void writeExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ch == null)
                {
                    throw new Exception($@"Перед записью необходимо считать данные!");
                }
                string name = "excelFile.xlsx";
                string path = Path.Combine(Directory.GetCurrentDirectory(), name);
                WriteHelper.WriteExcel(ch.Items, path);

                MessageBox.Show($@"Данные успешно записаны в файл: {name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if ((bool)overwrite.IsChecked)
                {
                    Global.overwritingFile = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }

        private void overwrite_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(bool)overwrite.IsChecked)
                {
                    Global.overwritingFile = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"Возникла проблема в работе приложения: {ex.Message}");
            }
        }
    }
}
