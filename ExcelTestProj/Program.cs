using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTestProj
{
    class Program
    {
        static void Main(string[] args)
        {
            OpenFile();
        }
        private static void OpenFile()
        {
            Excel excel = new Excel(@"D:/Heshvan2.xlsx", 1);
            int row = 1; int col = 0;

            Dictionary<string, Zman_Day> dictionary = new Dictionary<string, Zman_Day>();
            List<Zman_Day> days = new List<Zman_Day>();

            for(int i=1; i<excel.rw; i++)
            {
                Zman_Day day = new Zman_Day();
                day.Yom = excel.ReadCell(i, 0);
                day.Yom_Be_Shavua = excel.ReadCell(i, 1);
                day.Loazi_Date = excel.ReadCell(i, 2);
                day.Alot_Ashachar = excel.ReadCell(i, 3);
                day.Zman_talit = excel.ReadCell(i, 4);
                day.Netz = excel.ReadCell(i, 5);
                day.Netz_Mishur = excel.ReadCell(i, 6);
                day.Sof_Ks_Magen = excel.ReadCell(i, 7);
                day.Sof_Ks_Gra = excel.ReadCell(i, 8);
                day.Sof_Tfila_Magen = excel.ReadCell(i, 9);
                day.Sof_Tfila_Gra = excel.ReadCell(i, 10);
                day.Hatzot = excel.ReadCell(i, 11);
                day.Mincha_Gdola = excel.ReadCell(i, 12);
                day.Mincha_Ktana = excel.ReadCell(i, 13);
                day.Plag = excel.ReadCell(i, 14);
                day.Sunset = excel.ReadCell(i, 15);
                day.Tzet_Kochavim = excel.ReadCell(i, 16);
                day.Tzet_Kochavim_Tam = excel.ReadCell(i, 17);
                day.Daf_Yomi_Bavel = excel.ReadCell(i, 18);
                day.Daf_Yomi_yeru = excel.ReadCell(i, 19);
                day.Rambam_Yomi = excel.ReadCell(i, 20);
                day.Rambam_g_Prakim = excel.ReadCell(i, 21);
                Console.WriteLine("Read Day " + day.Loazi_Date);
                dictionary.Add(day.Loazi_Date, day);
                //days.Add(day);
            }
            excel.Destroy();
            Console.ReadKey();
            


        }

    }
}
