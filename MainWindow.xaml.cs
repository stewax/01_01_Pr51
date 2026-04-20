using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using Word_kazakov.Context;


namespace Word_kazakov
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadRooms();
        }
        public void LoadRooms()
        {
            for (int i = 1; i < 20; i++)
                Parent.Children.Add(new Elements.Room(i));
        }

        private void Report(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Files (*.docx)|*.docx";
            sfd.ShowDialog();
            if (sfd.FileName != "")
                OwnerContext.Report(sfd.FileName);
        }


        private void ReportPDF(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "PDF Files (*.pdf)|*.pdf";
            sfd.DefaultExt = "pdf";
            sfd.FileName = "Отчет_Жильцы";

            if (sfd.ShowDialog() == true)
            {
                string filePath = sfd.FileName;

                try
                {
                    OwnerContext.ReportPDF(filePath);

                    if (System.IO.File.Exists(filePath))
                    {
                        System.Threading.Thread.Sleep(100);

                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = filePath,
                            UseShellExecute = true
                        });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка: {ex.Message}");
                }
            }
        }
    }
}
