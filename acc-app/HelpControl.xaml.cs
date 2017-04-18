using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
    /// Logique d'interaction pour HelpControl.xaml
    /// </summary>
    public partial class HelpControl : UserControl
    {
        public HelpControl()
        {
            InitializeComponent();

            this.WebBrowser.NavigateToStream(Assembly.GetExecutingAssembly().GetManifestResourceStream("acc_app.Aide.html"));
        }
    }
}
