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

namespace ExcelUtilities
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

        private void btnMatchF741_Click(object sender, RoutedEventArgs e)
        {
            Forms.MatchF741 wF741 = new Forms.MatchF741();
            wF741.Show();
        }

        private void btnExceltoText_Click(object sender, RoutedEventArgs e)
        {
            Forms.ExcelToText wXlstoText = new Forms.ExcelToText();
            wXlstoText.Show();
        }
    }
}
