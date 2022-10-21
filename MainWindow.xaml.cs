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
using SpreadsheetLight;
using Microsoft.Win32;

namespace Name
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonAddName_Click(object sender, RoutedEventArgs e)
        {
            // check that textbox contains name and that this name isnt already in listbox
            if (!string.IsNullOrWhiteSpace(txtName.Text) && !lstNames.Items.Contains(txtName.Text))
            {
                // add name to listbox
                lstNames.Items.Add(txtName.Text);
                // creal textbox
                txtName.Clear();
            }
        }

        private void ButtonCreateExcel_Click(object sender, RoutedEventArgs e)
        {
            if(!lstNames.HasItems)
            {
                MessageBox.Show("Listbox has no items, add some", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveFileDialog saveFile = new SaveFileDialog();
            // file extension
            saveFile.Filter = "Excel | *.xlsx";
            saveFile.DefaultExt = "xlsx";

            if (saveFile.ShowDialog() == true)
            {
                CreateFile(saveFile.FileName);
            }
        }

        public void CreateFile(string path)
        {
            if(!string.IsNullOrWhiteSpace(path))
            {
                using (SLDocument sl = new SLDocument())
                {
                    // set first cell as Names
                    sl.SetCellValue("A1", "Names");

                    for(int i = 0; i < lstNames.Items.Count; i++)
                    {
                        sl.SetCellValue($"A{i + 2}", $"{lstNames.Items.GetItemAt(i)}");
                    }

                    // save excel file
                    sl.SaveAs(path);
                }
            }
            else
            {
                MessageBox.Show("Invalid path", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
