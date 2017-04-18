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
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace acc_app
{
    /// <summary>
    /// Logique d'interaction pour AddRecetteControl.xaml
    /// </summary>
    public partial class AddRecetteControl : UserControl
    {
        public delegate void AddRecetteButtonClickHandler(object sender, RoutedEventArgs e);
        public event AddRecetteButtonClickHandler AddRecetteButtonClicked;

        public AddRecetteControl()
        {
            InitializeComponent();
        }

        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            AddRecetteButtonClicked?.Invoke(sender, e);
        }
    }
}
