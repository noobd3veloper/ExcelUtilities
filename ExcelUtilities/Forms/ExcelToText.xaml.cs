using ExcelUtilities.Model;
using ExcelUtilities.Utilities;
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
using System.Windows.Shapes;
using org.apache.poi.openxml4j.opc;
using org.apache.poi.ss.usermodel;
using org.apache.poi.xssf.usermodel;
using org.apache.poi.hssf.usermodel;

namespace ExcelUtilities.Forms
{
    /// <summary>
    /// Interaction logic for ExcelToText.xaml
    /// </summary>
    public partial class ExcelToText : Window
    {
        private class ListViewItemList
        {
            public String inFile { get; set; }
            public String ouFile { get; set; }

        }

        public ExcelToText()
        {
            InitializeComponent();
        }

        private Dictionary<String, F820> hF820;
        private List<F820> lstF820;
        private void listView_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
        }

        private void listView_Drop(object sender, DragEventArgs e)
        {
            String[] files = (String[])e.Data.GetData(DataFormats.FileDrop);
            foreach(string file in files)
            {
                listView.Items.Add(new ListViewItemList { inFile = file, ouFile = "" });
            }
            
        }

        private void btnLoadFiles_Click(object sender, RoutedEventArgs e)
        {
            LoadFiles loadFiles = new LoadFiles();
            loadFiles.loadF820(txtF820.Text, out lstF820, out hF820);
            MessageBox.Show("Files Loaded");
        }

        private void btnF820_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "All Files | *.*";
            ofd.RestoreDirectory = true;
            ofd.Multiselect = false;
            Button b = (Button)sender;
            if (ofd.ShowDialog() == true)
            {
                if (b.Name.Equals("btnF820"))
                {
                    txtF820.Text = ofd.FileName;
                }
                if (b.Name.Equals("btnF741"))
                {
                    txtF741.Text = ofd.FileName;
                }
            }
        }


        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            String workCode;
            List<F741> f741s = new List<F741>();
            List<F741> outF741s = null;
            ExcelFilesProcessing xfp = new ExcelFilesProcessing();
            foreach (ListViewItemList o in listView.Items)
            {
                XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(o.inFile));
                for(int sheetCount = 0;sheetCount < workbook.getNumberOfSheets(); sheetCount++)
                {
                    XSSFSheet sheet = workbook.getSheetAt(sheetCount);
                    F820 f820 = lstF820.Find(x => x.Description.Trim().Equals(sheet.getSheetName().Trim()));
                    if (f820 != null)
                    {
                        workCode = f820.MntcWork;
                    }
                    if (!sheet.getSheetName().ToUpper().Contains("ESC") &&
                        !sheet.getSheetName().ToUpper().Contains("LIFT"))
                    {
                        xfp.CritNonCritNonLiftWorks(sheet, lstF820, out outF741s);
                    }
                    if (sheet.getSheetName().ToUpper().Contains("LIFT"))
                    {
                        xfp.CritNonCritLiftWorks(sheet, lstF820, out outF741s);
                    }
                    f741s.AddRange(outF741s);
                }
                workbook = null;
                WriteFiles wf = new WriteFiles();
                wf.WriteF741(f741s, txtF741.Text, pb1);
                
            }
        }
    }

    
}
