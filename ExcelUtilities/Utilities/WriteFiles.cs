using ExcelUtilities.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ExcelUtilities.Utilities
{
    class WriteFiles
    {
        public void WriteF741(List<F741> lstF741, String path, ProgressBar pb1)
        {

            StreamWriter sw = new StreamWriter(path, File.Exists(path));
            pb1.Maximum = lstF741.Count();
            pb1.Value = 0;
            foreach(F741 f741 in lstF741)
            {

                ExcelUtilities.App.DoEvents();
                pb1.Value++;
                sw.WriteLine(formatF741Object(f741));
                
            }
            sw.Flush();
            sw.Close();
        }

        private String formatF741Object(F741 f741)
        {
            String line =
                f741.BldngGl.PadRight(5) +
                f741.MntcWork.PadRight(3) +
                f741.MntcLabel.PadRight(2) +
                f741.LastHistory.PadRight(6) +
                f741.LastDone.PadRight(6) +
                f741.NextPlan.PadRight(6) +
                f741.TextRemarks.PadRight(120) +
                f741.WorkStatus.PadRight(1) +
                f741.LastDateInspect.PadRight(6);
            return line;

            
        }
    }
}
