using ExcelUtilities.Model;
using ExcelUtilities.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace ExcelUtilities.Forms
{
    /// <summary>
    /// Interaction logic for MatchF741.xaml
    /// </summary>
    public partial class MatchF741 : Window
    {

        private List<F741> lstNewF741;
        private List<F741> lstOldF741;
        private List<F742> lstF742;
        private List<F820> lstF820;
        private List<F800> lstF800;
        private List<F741> lstTurnOnDate;
        private Dictionary<String, F741> hOldF741;
        private Dictionary<String, F741> hNewF741;
        private Dictionary<String, F742> hF742;
        private Dictionary<String, F800> hF800;
        private Dictionary<String, F820> hF820;
        private Dictionary<String, F741> hTurnOnDate;

        public MatchF741()
        {
            InitializeComponent();
            lstNewF741 = new List<F741>();
            lstOldF741 = new List<F741>();
            lstTurnOnDate = new List<F741>();
            lstF742 = new List<F742>();
            lstF820 = new List<F820>();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "All Files | *.*";
            ofd.RestoreDirectory = true;
            ofd.Multiselect = false;
            Button b = (Button)sender;
            if (ofd.ShowDialog() == true)
            {
                if (b.Name.Equals("btnBrowseOldF741"))
                {
                    txtOldF741.Text = ofd.FileName;
                }
                if (b.Name.Equals("btnBrowseNewF741"))
                {
                    txtNewF741.Text = ofd.FileName;
                }
                if (b.Name.Equals("btnBrowseF820"))
                {
                    txtF820.Text = ofd.FileName;
                }
                if (b.Name.Equals("btnTurnOnDate"))
                {
                    txtTurnOnDate.Text = ofd.FileName;
                }
                if (b.Name.Equals("btnOutput"))
                {
                    txtOutput.Text = ofd.FileName;
                }
            }
        }

        private void btnLoad_Click(object sender, RoutedEventArgs e)
        {
            LoadFiles loadFiles = new LoadFiles();
            loadFiles.loadF741(txtOldF741.Text, out lstOldF741, out hOldF741);
            loadFiles.loadF820(txtF820.Text, out lstF820, out hF820);
            loadFiles.loadF741(txtNewF741.Text, out lstNewF741, out hNewF741);
            loadFiles.loadF741(txtTurnOnDate.Text, out lstTurnOnDate, out hTurnOnDate);
            MessageBox.Show("Files Loaded");
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {

            if (lstOldF741.Count == 0 ||
                lstF820.Count == 0 ||
                lstNewF741.Count == 0 ||
                lstTurnOnDate.Count == 0)
            {
                MessageBox.Show("Please load input files.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            gbProgrss.Visibility = Visibility.Visible;
            pb1.Maximum = lstOldF741.Count;
            Object countLock = new Object();
            String key;
            F741 outF741 = null, newF741;
            List<F741> lstOutputF741 = new List<F741>();
            List<F741> lstWriteF741 = new List<F741>();
            Int32 tmpNextSched;
            List<String> liftWorks = new List<String> { "10A", "100", "101", "102", "103",
                "104", "105", "106", "109", "110", "111"};
            String keyTurnedOn = "";
            foreach(F741 f741 in lstOldF741)
            {
                ExcelUtilities.App.DoEvents();

                //pb1.Dispatcher.Invoke(DispatcherPriority.Normal,
                //     new Action(() =>
                //     {
                //         ExcelUtilities.App.DoEvents();
                //         pb1.Value++;
                //         gbProgrss.Header = String.Format("Processing {0} of {1}", pb1.Value, pb1.Maximum);
                //     }));
                //ExcelUtilities.App.DoEvents();
                pb1.Value++;
                gbProgrss.Header = String.Format("Processing {0} of {1}", pb1.Value, pb1.Maximum);
                key = f741.BldngGl + f741.MntcWork + f741.MntcLabel;
                F741 value;

                if (hNewF741.TryGetValue(key, out value))
                {
                    newF741 = value;
                }
                else
                {
                    newF741 = null;
                }

                if (newF741 != null)
                {

                    if (liftWorks.Contains(newF741.MntcWork))
                    {
                        keyTurnedOn = f741.BldngGl + f741.MntcLabel;
                        outF741 = activateLift(keyTurnedOn, f741, newF741);
                        //Lift
                    }
                    else
                    {
                        //Non Lift
                        if (!f741.LastDone.Trim().Equals(""))
                        {
                            if (f741.LastDone.Trim().Equals(""))
                            {
                                f741.LastDone = "000101";
                            }
                            if (Int32.Parse(newF741.LastDone.Substring(0, 4)) >= Int32.Parse(f741.LastDone.Substring(0, 4)))
                            {
                                if (newF741.LastDone == "000101")
                                {
                                    outF741 = newF741;
                                    outF741.WorkStatus = "Q";
                                }
                                else
                                {
                                    outF741 = newF741;
                                    F820 f820 = lstF820.Find(x => x.MntcWork == outF741.MntcWork);
                                    tmpNextSched = Int32.Parse(newF741.LastDone) +
                                        (Int32.Parse(f820.CycleYear) * 100);
                                    outF741.LastHistory = f741.LastDone.ToString();
                                    outF741.NextPlan = tmpNextSched.ToString("000000");
                                }
                            }
                            else if (newF741.LastDone == "000101")
                            {
                                outF741 = newF741;
                                outF741.WorkStatus = "Q";
                            }
                            else if (Int32.Parse(newF741.LastDone.Substring(0, 4)) <= Int32.Parse(f741.LastDone.Substring(0, 4)))
                            {
                                outF741 = newF741;
                                outF741.LastDateInspect = "      ";
                            }
                        }
                        else
                        {
                            outF741 = f741;
                            outF741.LastDateInspect = "";
                        }
                    }
                }
                else
                {
                    outF741 = f741;
                    outF741.LastDateInspect = "      ";
                }
                lstOutputF741.Add(outF741);
            }

            maxBatteryDate(lstOutputF741, out lstWriteF741);
            WriteFiles writeFiles = new WriteFiles();
            writeFiles.WriteF741(lstWriteF741, txtOutput.Text, pb1);
        }


        private F741 activateLift(String key, F741 f741, F741 newF741)
        {
            F741 liftF741 = new F741();
            F741 value;
            F741 f7, f7Old;
            Int32 tmpSched = 0;
            String keyOldF741 = f741.BldngGl + "10A" + f741.MntcLabel;
            if(hTurnOnDate.TryGetValue(key, out value))
            {
                f7 = value;
            }else
            {
                f7 = null;
            }
            if (hOldF741.TryGetValue(keyOldF741, out value))
            {
                f7Old = value;
            }else
            {
                f7Old = null;
            }

            if (f7 == null)
            {

                if (newF741.LastDone.Trim().Equals("000101") && !(newF741.MntcWork.Trim().Equals("104")))
                {
                    liftF741 = newF741;
                    liftF741.LastHistory = f741.LastDone;
                    liftF741.WorkStatus = "Q";
                }else if (Int32.Parse(newF741.LastDone.Substring(0,4)) > Int32.Parse(f741.LastDone.Substring(0,4)))
                {
                    if (Int32.Parse(newF741.LastDone.Substring(0, 6)) > Int32.Parse(dtpUpdate.Text))
                    {
                        liftF741 = f741;
                    }
                    else
                    {
                        liftF741 = newF741;
                        F820 f820 = lstF820.Find(f8 => f8.MntcWork == liftF741.MntcWork);
                        tmpSched = Int32.Parse(newF741.LastDone) + (Int32.Parse(f820.CycleYear) * 100);
                        liftF741.LastHistory = f741.LastDone.ToString();
                        liftF741.NextPlan = tmpSched.ToString("000000");
                    }
                }else
                {
                    liftF741 = f741;
                }

            }else if (f7 != null && f7Old == null)
            {
                liftF741 = f741;
            }else if(f7 != null && f7Old != null)
            {
                if (!f7.LastDone.Trim().Equals("000101"))
                {
                    if(newF741.LastDone.Trim().Equals("000101") && !newF741.LastDone.Trim().Equals("104"))
                    {
                        liftF741 = newF741;
                        liftF741.LastHistory = f741.LastDone;
                        liftF741.WorkStatus = "Q";
                    }
                    else if (Int32.Parse(f7.LastDone.Substring(0, 4)) > Int32.Parse(f741.LastDone.Substring(0, 4)))
                    {
                        if (Int32.Parse(f7.LastDone.Substring(0, 6)) > Int32.Parse(dtpUpdate.Text))
                        {
                            liftF741 = f741;
                        }
                        else
                        {
                            liftF741 = new F741();
                            liftF741.BldngGl = f7.BldngGl;
                            liftF741.WorkStatus = "A";
                            liftF741.TextRemarks = f7.TextRemarks;
                            liftF741.MntcLabel = f7.MntcLabel;
                            liftF741.MntcWork = f7.MntcWork;
                            F820 f820 = lstF820.Find(f8 => f8.MntcWork == liftF741.MntcWork);
                            tmpSched = Int32.Parse(f7.LastDone) + (Int32.Parse(f820.CycleYear) * 100);
                            liftF741.LastDone = f7.LastDone;
                            liftF741.LastHistory = f741.LastDone.ToString();
                            liftF741.NextPlan = tmpSched.ToString("000000");
                        }
                    }
                    else if (Int32.Parse(newF741.LastDone.Substring(0, 4)) > Int32.Parse(f741.LastDone.Substring(0, 4)))
                    {
                        if (Int32.Parse(newF741.LastDone.Substring(0, 6)) > Int32.Parse(dtpUpdate.Text))
                        {
                            liftF741 = f741;
                        }
                        else
                        {
                            liftF741 = newF741;
                            F820 f820 = lstF820.Find(f8 => f8.MntcWork == liftF741.MntcWork);
                            tmpSched = Int32.Parse(newF741.LastDone) + (Int32.Parse(f820.CycleYear) * 100);
                            liftF741.LastHistory = f741.LastDone.ToString();
                            liftF741.NextPlan = tmpSched.ToString("000000");
                        }
                    }
                }else
                {
                    liftF741 = newF741;
                }
            }
            return liftF741;
        }

        private void maxBatteryDate(List<F741> lstInputF741, out List<F741> lstOutputF741)
        {
            lstOutputF741 = new List<F741>();
            List<String> lstBuildingGl = lstInputF741.Select(x => x.BldngGl).Distinct().ToList();
            Dictionary<String, F741> dictOutputF741 = lstInputF741.GroupBy(f => f.BldngGl + f.MntcWork + f.MntcLabel)
                .ToDictionary(x => x.Key, x => x.First());
            F741 outValue;
            String[] batteryArdWorks = new String[] { "101", "102", "110" };
            String[] batteryEbopWorks = new String[] { "103", "111"};
            int[] dates;
            int maxDate;
            Parallel.ForEach(lstBuildingGl, gl =>
            {
                if(dictOutputF741.TryGetValue(gl.Substring(0,5) + "101" + gl.Substring(5,2), out outValue))
                {
                    dates = new int[batteryArdWorks.Count()];
                    for(int i = 0; i< batteryArdWorks.Count(); i++)
                    {
                        dates[i] = int.Parse(dictOutputF741[gl.Substring(0, 5) + batteryArdWorks[i] + gl.Substring(5, 2)].LastDone);
                    }

                    maxDate = dates.Max();
                    foreach(int i in Array.FindAll(dates, x => x != maxDate))
                    {
                        dictOutputF741[gl.Substring(0, 5) + batteryArdWorks[i] + gl.Substring(5, 2)].LastDone = "000101";
                        dictOutputF741[gl.Substring(0, 5) + batteryArdWorks[i] + gl.Substring(5, 2)].WorkStatus = "Q";
                    }
                    
                }

                if (dictOutputF741.TryGetValue(gl.Substring(0, 5) + "103" + gl.Substring(5, 2), out outValue))
                {
                    dates = new int[batteryEbopWorks.Count()];
                    for (int i = 0; i < batteryArdWorks.Count(); i++)
                    {
                        dates[i] = int.Parse(dictOutputF741[gl.Substring(0, 5) + batteryEbopWorks[i] + gl.Substring(5, 2)].LastDone);
                    }

                    maxDate = dates.Max();
                    foreach (int i in Array.FindAll(dates, x => x != maxDate))
                    {
                        dictOutputF741[gl.Substring(0, 5) + batteryEbopWorks[i] + gl.Substring(5, 2)].LastDone = "000101";
                        dictOutputF741[gl.Substring(0, 5) + batteryEbopWorks[i] + gl.Substring(5, 2)].WorkStatus = "Q";
                    }

                }
            });
            lstOutputF741 = dictOutputF741.Values.ToList();
        }

        private void test_Click(object sender, RoutedEventArgs e)
        {
            int[] dates = new int[] { 3, 2, 3 };
            int maxDate;
            maxDate = dates.Max();
            MessageBox.Show(maxDate.ToString());
            MessageBox.Show(Array.FindAll(dates,x => x == maxDate).Count().ToString());

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            gbProgrss.Visibility = Visibility.Hidden;
        }
    }
}
