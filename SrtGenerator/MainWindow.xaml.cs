using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Diagnostics;
using SrtGenerator.Models;
using SrtGenerator.Serialization;

namespace SrtGenerator
{ 

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            parseButton.IsEnabled = false;
            generateButton.IsEnabled = false;

            listBox.PreviewMouseDown += ListBox_PreviewMouseDown;
        }

        private void ListBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement(listBox, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (item != null)
            {
                try
                {
                    Process.Start("notepad.exe", item.Content.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erreur");
                }
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();            
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Document (.xlsx)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result != true)
            {
                return;
            }

            string filename = dlg.FileName;
            textBox.Text = filename;

            parseButton.IsEnabled = true;
            


        }

        SrtModel currentSrt;
   
        private void Parse_Click(object sender, RoutedEventArgs e)
        {  

            try
            {
                generateButton.IsEnabled = false;

                currentSrt = XlsParser.Load(textBox.Text);   

                generateButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }

        }



        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            if(currentSrt==null)
            {
                return;
            }

            listBox.Items.Clear();

            try
            {

                foreach (var language in currentSrt.Languages)
                {
                    var result = currentSrt.GetSrtText(language);

                    String pathToFile = currentSrt.Path + "\\" + currentSrt.Filename + "." + language + ".srt";

                    System.IO.File.WriteAllText(pathToFile, result);

                    listBox.Items.Add(pathToFile);
                    
                }


                Process.Start(currentSrt.Path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erreur");
            }

        }     

     
   
    }
}
