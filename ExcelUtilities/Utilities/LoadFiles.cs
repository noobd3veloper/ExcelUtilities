using ExcelUtilities.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelUtilities.Utilities
{
    class LoadFiles
    {

        String[] notValidString = new String[] { "\u001a", "", "\0\0\0\t\v\0\0\0\t\f\0\0\0\t", "\u0001\0\0\t\v\u0001\0\0\t\f\u0001\0\0\t","\u0002\0\0\t\v\u0002\0\0\t\f\u0002\0\0\t" };
        public void loadF741(string path, out List<F741> lstF741)
        {
            Dictionary<String, F741> hF741;
            this.loadF741(path, out hF741);
            lstF741 = hF741.Values.ToList();
        }

        public void loadF741(string path, out List<F741> lstF741, out Dictionary<String, F741> hF741)
        {
            this.loadF741(path, out hF741);
            lstF741 = hF741.Values.ToList();
        }


        public void loadF741(string path, out Dictionary<String, F741> hF741)
        {
            hF741 = new Dictionary<string, F741>();
            String[] lines = File.ReadAllLines(path);
            String key;
            Int32 recordCount = 0;
            foreach (String item in lines)
            {
                recordCount++;
                F741 f = new F741();
                if (Array.FindAll(notValidString, x => x==item).Count() == 0)
                {
                    f = defineF741Layout(item);
                    key = f.BldngGl + f.MntcWork + f.MntcLabel;
                    if (!hF741.ContainsKey(key))
                    {
                        hF741.Add(key, f);
                    }
                    else
                    {
                        if (Int32.Parse(hF741[key].LastDone) >= Int32.Parse(f.LastDone))
                        {
                            Console.WriteLine("F741 Contains" + key);
                        }
                        else
                        {
                            hF741[key] = f;
                        }
                    }
                }
            }
            Console.WriteLine("F741 Contains {0} records", recordCount);
        }


        public void loadF820(string path, out List<F820> lstF820)
        {
            Dictionary<String, F820> hF820;
            this.loadF820(path, out hF820);
            lstF820 = hF820.Values.ToList();
        }

        public void loadF820(string path, out List<F820> lstF820, out Dictionary<String, F820> hF820)
        {
            this.loadF820(path, out hF820);
            lstF820 = hF820.Values.ToList();
        }

        public void loadF820(string path, out Dictionary<String, F820> hF820)
        {
            hF820 = new Dictionary<string, F820>();
            String[] lines = File.ReadAllLines(path);
            String key;
            Int32 recordCount = 0;
            foreach (String item in lines)
            {
                recordCount++;
                F820 f = new F820();
                if (Array.FindAll(notValidString, x => x == item).Count() == 0)
                {
                    f = defineF820Layout(item);
                    key = f.MntcWork;
                    if (!hF820.ContainsKey(key))
                    {
                        hF820.Add(key, f);
                    }
                    else
                    {
                         Console.WriteLine("F820 Contains" + key);
                    }
                }
            }
            Console.WriteLine("F741 Contains {0} records", recordCount);
        }
        private F741 defineF741Layout(String record)
        {
            F741 f741 = new F741();
            f741.BldngGl = record.Substring(0, 5);
            f741.MntcWork = record.Substring(5, 3);
            f741.MntcLabel = record.Substring(8, 2);
            f741.LastHistory = record.Substring(10, 6);
            DateTime x;
            CultureInfo culture = CultureInfo.InvariantCulture;
            DateTimeStyles ds;
            String[] str = { "yyyyMMdd" };
            ds = DateTimeStyles.None;
            if (record.Substring(16, 6).Trim().Equals(String.Empty))
            {
                f741.LastDone = record.Substring(16, 6);
            }else if(DateTime.TryParseExact(record.Substring(16, 6) + "01",str,culture,ds,out x))
            {
                f741.LastDone = record.Substring(16, 6);
            }
            else
            {
                MessageBox.Show("Error : " + record);
            }
            f741.NextPlan = record.Substring(22, 6);
            f741.TextRemarks = record.Substring(28, 120);
            f741.WorkStatus = record.Substring(148, 1);
            f741.LastDateInspect = record.Substring(149, 6);
            return f741;
        }

        private F820 defineF820Layout(String record)
        {
            F820 f820 = new F820();
            f820.MntcWork = record.Substring(0, 3);
            f820.SubType = record.Substring(3, 2);
            f820.Description = record.Substring(5, 60);
            f820.CycleYear = record.Substring(65, 2);
            f820.Category = record.Substring(67, 5);
            return f820;
        }
    }
}
