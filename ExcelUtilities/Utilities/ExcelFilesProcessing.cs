using ExcelUtilities.Model;
using org.apache.poi.hssf.usermodel;
using org.apache.poi.ss.usermodel;
using org.apache.poi.xssf.usermodel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUtilities.Utilities
{
    class ExcelFilesProcessing
    {
        private String lastDone, dateInspect, works, workCode, label;
        private DateTime result;
        public void CritNonCritNonLiftWorks(XSSFSheet sheet,List<F820> F820s, out List<F741> F741s)
        {
            F741s = new List<F741>();
            F741 f741;
            Cell cell = null;
            
            foreach(Row row in sheet)
            {
                f741 = new F741();
                cell = row.getCell(2);
                if (String.IsNullOrEmpty(cell.ToString()))
                    continue;

                f741.BldngGl = cell.ToString().Trim();
                cell = row.getCell(9); //Last Done Date
                lastDone = cell.ToString().Trim();
                
                DateTime.TryParse(lastDone, out result);
                if (result == null)
                {
                    lastDone = "000101";
                }else
                {
                    if((!lastDone.Equals("/") && !String.IsNullOrEmpty(lastDone)) &&
                            result.Year.ToString("0000") == "0001")
                    {
                        continue;
                    }
                    lastDone = result.Year.ToString("0000") + result.Month.ToString("00");
                }
                f741.WorkStatus = "A";
                if (lastDone.Equals("000101"))
                {
                    f741.WorkStatus = "Q";
                }
                f741.LastDone = lastDone;
                cell = row.getCell(13); //Text Remarks
                if (String.IsNullOrEmpty(cell.ToString()))
                {
                    f741.TextRemarks = String.Empty;
                }else
                {
                    if (cell.ToString().Trim().Length >= 118)
                    {
                        f741.TextRemarks = cell.ToString().Trim().Substring(0, 115);
                    }else
                    {
                        f741.TextRemarks = cell.ToString().Trim();
                    }
                }
                cell = row.getCell(14); //Last Date Inspect
                if (!String.IsNullOrEmpty(cell.ToString()))
                {

                    dateInspect = cell.ToString().Trim();
                    DateTime.TryParse(dateInspect, out result);
                    if (result == null)
                    {
                        dateInspect = "";
                    }else
                    {
                        dateInspect = result.Year.ToString("0000") + result.Month.ToString("00");
                    }
                }else
                {
                    dateInspect = "";
                }
                f741.LastDateInspect = dateInspect;
                cell = row.getCell(8);
                works = cell.ToString().Trim();
                F820 f820 = F820s.Find(x => x.Description.Trim().Equals(works.Replace("FPS  ","FPS ")));
                workCode = f820.MntcWork;
                f741.MntcWork = workCode;
                if (f741.MntcWork.Equals("005") || f741.MntcWork.Equals("007"))
                {
                    f741.MntcLabel = "1";
                }else
                {
                    f741.MntcLabel = "";
                }
                f741.LastHistory = "";
                f741.NextPlan = "";
                F741s.Add(f741);    
            }
           
        }

        public void CritNonCritLiftWorks(XSSFSheet sheet, List<F820> F820s, out List<F741> F741s)
        {
            F741s = new List<F741>();
            F741 f741;
            Cell cell = null;
            String[] workArray;
            foreach (Row row in sheet)
            {
                f741 = new F741();
                cell = row.getCell(2);
                if (String.IsNullOrEmpty(cell.ToString()))
                    continue;
                f741.BldngGl = cell.ToString().Trim();
                cell = row.getCell(10); //Last Done Date
                lastDone = cell.ToString().Trim();

                DateTime.TryParse(lastDone, out result);
                if (result == null)
                {
                    lastDone = "000101";
                }
                else
                {
                    if ((!lastDone.Equals("/") && !String.IsNullOrEmpty(lastDone)) &&
                            result.Year.ToString("0000") == "0001")
                    {
                        continue;
                    }
                    lastDone = result.Year.ToString("0000") + result.Month.ToString("00");
                }
                f741.WorkStatus = "A";
                if (lastDone.Equals("000101"))
                {
                    f741.WorkStatus = "Q";
                }
                f741.LastDone = lastDone;
                cell = row.getCell(14); //Text Remarks
                if (String.IsNullOrEmpty(cell.ToString()))
                {
                    f741.TextRemarks = String.Empty;
                }
                else
                {
                    if (cell.ToString().Trim().Length >= 118)
                    {
                        f741.TextRemarks = cell.ToString().Trim().Substring(0, 115);
                    }
                    else
                    {
                        f741.TextRemarks = cell.ToString().Trim();
                    }
                }
                cell = row.getCell(8);
                workArray = cell.ToString().Split(new char[] { '-' });
                label = workArray[0].ToString().Substring(5, 2);

                if (workArray.Length > 2)
                {
                    works = "LIFT - " + workArray[1].ToString().Trim() + "-" + workArray[2].ToString().Trim();
                }else
                {
                    works = "LIFT - " + workArray[1].ToString().Trim();
                }

                cell = row.getCell(15); //Last Date Inspect
                if (!String.IsNullOrEmpty(cell.ToString()))
                {

                    dateInspect = cell.ToString().Trim();
                    DateTime.TryParse(dateInspect, out result);
                    if (result == null)
                    {
                        dateInspect = "";
                    }
                    else
                    {
                        dateInspect = result.Year.ToString("0000") + result.Month.ToString("00");
                    }
                }
                else
                {
                    dateInspect = "";
                }
                f741.LastDateInspect = dateInspect;

                F820 f820 = F820s.Find(x => x.Description.Trim().Equals(works));
                workCode = f820.MntcWork;
                f741.MntcWork = workCode;
                f741.MntcLabel = label;
                f741.LastHistory = "";
                f741.NextPlan = "";
                F741s.Add(f741);
                //Test
            }
        }

        public void CritNonCritEscalatorWorks(XSSFSheet sheet, List<F820> F820s, out List<F741> F741s)
        {
            F741s = new List<F741>();
            F741 f741;
            Cell cell = null;
            String[] workArray;
            foreach (Row row in sheet)
            {
                f741 = new F741();
                cell = row.getCell(2);
                if (String.IsNullOrEmpty(cell.ToString()))
                    continue;
                f741.BldngGl = cell.ToString().Trim();
                cell = row.getCell(9); //Last Done Date
                lastDone = cell.ToString().Trim();

                DateTime.TryParse(lastDone, out result);
                if (result == null)
                {
                    lastDone = "000101";
                }
                else
                {
                    if ((!lastDone.Equals("/") && !String.IsNullOrEmpty(lastDone)) &&
                            result.Year.ToString("0000") == "0001")
                    {
                        continue;
                    }
                    lastDone = result.Year.ToString("0000") + result.Month.ToString("00");
                }
                f741.WorkStatus = "A";
                if (lastDone.Equals("000101"))
                {
                    f741.WorkStatus = "Q";
                }
                f741.LastDone = lastDone;
                cell = row.getCell(13); //Text Remarks
                if (String.IsNullOrEmpty(cell.ToString()))
                {
                    f741.TextRemarks = String.Empty;
                }
                else
                {
                    if (cell.ToString().Trim().Length >= 118)
                    {
                        f741.TextRemarks = cell.ToString().Trim().Substring(0, 115);
                    }
                    else
                    {
                        f741.TextRemarks = cell.ToString().Trim();
                    }
                }
                cell = row.getCell(8); //ESC - Work
                workArray = cell.ToString().Split(new char[] { '-' });
                label = workArray[0].ToString().Substring(5, 2);
                works = "ESC - " + workArray[1].ToString().Trim();

                cell = row.getCell(14); //Last Date Inspect
                if (!String.IsNullOrEmpty(cell.ToString()))
                {

                    dateInspect = cell.ToString().Trim();
                    DateTime.TryParse(dateInspect, out result);
                    if (result == null)
                    {
                        dateInspect = "";
                    }
                    else
                    {
                        dateInspect = result.Year.ToString("0000") + result.Month.ToString("00");
                    }
                }
                else
                {
                    dateInspect = "";
                }
                f741.LastDateInspect = dateInspect;

                F820 f820 = F820s.Find(x => x.Description.Trim().Equals(works));
                workCode = f820.MntcWork;
                f741.MntcWork = workCode;
                f741.MntcLabel = label;
                f741.LastHistory = "";
                f741.NextPlan = "";
                F741s.Add(f741);
                //Test
            }
        }
    }
}
