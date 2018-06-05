using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelHighlight
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string FullPath = System.IO.Path.GetFullPath("Words.xml");

                //string xmlSource = "C:\\Users\\deell\\Desktop\\Excel\\ExcelHighlight\\Words.xml";
                //XDocument.Load(xmlSource);
                string[] arr = XDocument.Load(FullPath).Descendants("Highlight").Descendants().Select(x => x.ToString()).ToArray();
                //.Select(element => element.Value).ToArray();

                Application xlApp = new Application();
                string excelFile = @"‪‪D:\Test.xls";
                    //System.IO.Path.GetFullPath("TestCases.xlsx");
                Workbook xlWorkbook = (Workbook)xlApp.Workbooks.Open(excelFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Sheets sheet = xlWorkbook.Worksheets;
                string str;
                int rCnt = 0;
                int cCnt = 0;

                Worksheet xlWorkSheet4;
                Range range;
                xlWorkSheet4 = (Worksheet)sheet.get_Item(1);
                Range last3 = xlWorkSheet4.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                range = xlWorkSheet4.get_Range("A1", last3);
                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        if (range.Cells[rCnt, cCnt].Value2 is string)
                        {
                            str = (string)(range.Cells[rCnt, cCnt] as Range).Value2;
                            if (str == null)
                            {
                                Console.WriteLine("null");
                            }
                            else
                            {
                                str.Replace("\\", "");
                                string[] words = str.Split(' ');
                                foreach (string arrs in arr)
                                {
                                    foreach (string word in words)
                                    {
                                        if (word == arrs)
                                        {

                                            var cell = (range.Cells[rCnt, cCnt] as Range);

                                            cell.Font.Bold = 1;
                                            cell.Font.Color = ColorTranslator.ToOle(Color.Red);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("not string");
                        }
                    }
                }
            }
            finally
            {
               
            }
        }
    }
}
