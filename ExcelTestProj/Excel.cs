using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestProj
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        _Excel.Range range;
        public int rw = 0;
        public int cl = 0;
        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];

            range = ws.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            try
            {
                string x = (ws.Cells[i, j] as _Excel.Range).Value2.ToString();

                if (j == 3)
                {
                    long dateNum = long.Parse(ws.Cells[i, j].Value2.ToString());
                    DateTime result = DateTime.FromOADate(dateNum);
                    var ss = result.ToString("dd-MM-yyyy");
                    return ss;
                }
                else if (j > 3 && j <= 18)
                {
                    double dt = Convert.ToDouble((ws.Cells[i, j] as _Excel.Range).Value2.ToString());

                    int miliseconds = (int)Math.Round(dt * 86400000);
                    int hour = miliseconds / (60/*minutes*/* 60/*seconds*/* 1000);
                    miliseconds = miliseconds - hour * 60/*minutes*/* 60/*seconds*/* 1000;
                    int minutes = miliseconds / (60/*seconds*/* 1000);
                    miliseconds = miliseconds - minutes * 60/*seconds*/* 1000;
                    int seconds = miliseconds / 1000;

                    string fHour = hour < 10 ? "0" + hour.ToString() : hour.ToString();
                    string fMin = minutes < 10 ? "0" + minutes.ToString() : minutes.ToString();

                    return fHour + ":" + fMin;
                }

                return x;
            }
            catch (Exception e)
            {
                string sDate = (ws.Cells[i, j] as _Excel.Range).Value2.ToString();
                return sDate;
            }

        }

        public void Destroy()
        {
            wb = null;
            ws = null;
        }

        private DateTime ConvertToDateTime(double excelDate)
        {
            if (excelDate < 1)
            {
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            }
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            if (excelDate > 60d)
            {
                excelDate = excelDate - 2;
            }
            else
            {
                excelDate = excelDate - 1;
            }
            return dateOfReference.AddDays(excelDate);
        }
    }
}
