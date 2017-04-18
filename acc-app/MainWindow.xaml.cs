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

namespace acc_app
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel excel;

        public MainWindow()
        {
            InitializeComponent();

            this.@addRecette.IsEnabled = false;
            this.@addDepense.IsEnabled = false;

            this.ContentHolder.Content = new LogoControl();
        }

        private void MenuItemExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void MenuItemHelp_Click(object sender, RoutedEventArgs e)
        {
            this.ContentHolder.Content = new HelpControl();
        }

        private void MenuItemNew_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();

            dialog.Filter = "Excel file (*.xlsx)|*.xlsx";
            dialog.AddExtension = true;

            if (dialog.ShowDialog() == true)
            {
                this.excel = new Excel(dialog.FileName);

                try
                {
                    this.excel.Create();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(String.Format("Unable to create the Excel file ({0}).", exception.Message), "Erreur");
                }
                finally
                {
                    this.@new.IsEnabled = false;
                    this.@open.IsEnabled = false;

                    this.@addRecette.IsEnabled = true;
                    this.@addDepense.IsEnabled = true;
                }
            }

        }

        private void MenuItemOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "Excel file (*.xlsx)|*.xlsx";

            if (dialog.ShowDialog() == true)
            {
                this.excel = new Excel(dialog.FileName);

                try
                {
                    this.excel.Open();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(String.Format("Unable to open the Excel file ({0}).", exception.Message), "Erreur");
                }
                finally
                {
                    this.@new.IsEnabled = false;
                    this.@open.IsEnabled = false;

                    this.@addRecette.IsEnabled = true;
                    this.@addDepense.IsEnabled = true;
                }
            }
        }

        private void MenuItemAddRecette_Click(object sender, RoutedEventArgs e)
        {
            if (!this.excel.IsOpen)
            {
                MessageBox.Show("First, you have to create or open the Excel file.", "Erreur");

                return;
            }

            try
            {
                AddRecetteControl control = new AddRecetteControl();
                control.AddRecetteButtonClicked += new AddRecetteControl.AddRecetteButtonClickHandler(AddRecetteButton_Click);
                this.ContentHolder.Content = control;
            }
            catch (Exception exception)
            {
                MessageBox.Show(String.Format("Unable to add a recette ({0}).", exception.Message), "Erreur");
            }
        }

        private void MenuItemAddDepense_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddRecetteButton_Click(object sender, RoutedEventArgs e)
        {
            AddRecetteControl control = (AddRecetteControl)this.ContentHolder.Content;
            List<Object> values = new List<Object>();

            try
            {
                DatePicker controlFirst = (DatePicker)control.FindName("date");
                String date = controlFirst.SelectedDate.Value.ToString("dd/MM/yyyy");
                values.Add(date);

                ComboBox controlSecond = (ComboBox)control.FindName("mode");
                String mode = ((ComboBoxItem)controlSecond.SelectedItem).Content.ToString();
                values.Add(mode);

                TextBox controlThird = (TextBox)control.FindName("libelle");
                String libelle = controlThird.Text;
                if (String.IsNullOrEmpty(libelle.Trim())) throw new Exception();
                values.Add(libelle);

                ComboBox controlFourth = (ComboBox)control.FindName("provenance");
                String provenance = ((ComboBoxItem)controlFourth.SelectedItem).Content.ToString();
                values.Add(provenance);

                TextBox controlFifth = (TextBox)control.FindName("montant");
                String montant = controlFifth.Text;
                if (String.IsNullOrEmpty(montant.Trim())) throw new Exception();
                values.Add(Decimal.Parse(montant));
            }
            catch
            {
                values.Clear();
                MessageBox.Show("You have to set all items.", "Erreur");
            }

            try
            {
                this.excel.AddRecette(values);
            }
            catch (Exception exception)
            {
                MessageBox.Show(String.Format("Unable to add recette ({0}).", exception.Message), "Erreur");
            }
        }
    }
}
