using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelUtilities.Model
{
    [Serializable]
    class F741
    {
        public String BldngGl { get; set; }
        public String MntcWork { get; set; }
        public String MntcLabel { get; set; }
        public String LastHistory { get; set; }
        public String LastDone { get; set; }
        public String NextPlan { get; set; }
        public String TextRemarks { get; set; }
        public String WorkStatus { get; set; }
        public String LastDateInspect { get; set; }
    }
}
