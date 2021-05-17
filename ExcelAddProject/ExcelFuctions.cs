using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddProject
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first Excel-DNA function")]
        public static string MyFirstFunction(string name)
        {
            return "Hello " + name;
        }
        [ExcelFunction(Description = "Joins a string to a number", Category = "My functions")]
        public static string JoinThem([ExcelArgument(Description = "Input string",Name = "Word",AllowReference =true)] string str, [ExcelArgument(Description = "Input number", AllowReference = true)] double val)
        {
            return str + val;
        }
        [ExcelFunction(Description = "Multiplies two numbers", Category = "Useful functions")]
        public static double MultiplyThem(double x, double y)
        {
            return x * y;
        }
        [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
        public static double AddThem(
    [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")]
    double v1,
    [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]
    double v2)
        {
            return v1 + v2;
        }
    }
    public class RepairRates
    {
        public static List<Weld> WeldData()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "Wb сборка.xlsx") != null)
            {
                wb = xlApp.Workbooks["Wb сборка.xlsx"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\002-repair rates\Wb сборка.xlsx");
            }
            System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Хотите обновить WB сборку перед расчетом брака", "Repair Rate Counter 9000", System.Windows.Forms.MessageBoxButtons.YesNo);
            if (dialogResult == System.Windows.Forms.DialogResult.Yes)
            {
                ((Worksheet)wb.Worksheets["All"]).Activate();
                ((Worksheet)wb.Worksheets["All"]).AutoFilter.ShowAllData();
                ((Worksheet)wb.Worksheets["NEWS BASE"]).ListObjects["NEWS_BASE"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Workshop"]).ListObjects["Workshop"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Erection"]).ListObjects["Erection"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Flare"]).ListObjects["Flare"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["All"]).ListObjects["All"].QueryTable.Refresh(false);

                Range Selection = (((Worksheet)wb.Worksheets["All"]).ListObjects["All"].HeaderRowRange.Find("Resultat"));
                string Numberum = Selection.Address[false, false, XlReferenceStyle.xlA1, false];
                if ((((Worksheet)wb.Worksheets["All"]).ListObjects["All"].Range.Rows.Count - Selection.End[XlDirection.xlDown].Row) > 1)
                {


                    Selection = Selection.get_Offset(1, 0);
                    Selection = Selection.get_Resize(1, Selection.Column + Selection.End[XlDirection.xlToRight].Column);
                    Selection.Copy();
                        //[0, Selection.Column + Selection.End[XlDirection.xlToRight].Column].Copy();
                    Selection = Selection.get_Offset(Selection.End[XlDirection.xlDown].Row - Selection.Row, 0);
                    Selection = Selection.get_Resize(Selection.End[XlDirection.xlDown].Row-Selection.Row+1, Selection.Columns.Count);
                    Selection.Select();
                    
                    ((Worksheet)xlApp.ActiveSheet).Paste();
                    ((Worksheet)xlApp.ActiveSheet).Calculate();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Новых стыков не найдено");
                }
            }
            else if (dialogResult == System.Windows.Forms.DialogResult.No)
            {
                
            }
            Worksheet ws = wb.Worksheets["All"];
            ListObject WeldingBook = ws.ListObjects["All"];
            var WBarray = (object[,])WeldingBook.DataBodyRange.Value;
            List<Weld> WBook = new List<Weld>();
            //Workbook wb = xlApp.ActiveWorkbook;
            //Worksheet ws = wb.Worksheets[1];
            //ListObject WeldingBook = ws.ListObjects["All"];
            //WeldingBook.Range.Select();
            //List<Range> Base = WeldingBook.DataBodyRange.Rows.Cast<Range>().ToList();
            //var values = new object[WeldingBook.ListRows.Count, WeldingBook.ListColumns.Count];
            //values = WeldingBook.DataBodyRange.Value;

            //for (int i = 0; i < WeldingBook.ListRows.Count; i++)
            //{
            //    WBook.Add(new Weld());
            //    WBook[i].DrawingNum = WeldingBook.DataBodyRange[i + 1, 8].Value2;
            //    WBook[i].ISONum = WeldingBook.DataBodyRange[i + 1, 9].Value2;
            //    WBook[i].WeldNumber = WeldingBook.DataBodyRange[i + 1, 26].Value2;
            //    WBook[i].EndDate = WeldingBook.DataBodyRange[i + 1, 65].Value;
            //    WBook[i].WeldMaterial = WeldingBook.DataBodyRange[i + 1, 11].Value2;
            //    WBook[i].WeldDiam = WeldingBook.DataBodyRange[i + 1, 58].Value;
            //    WBook[i].WeldThick = WeldingBook.DataBodyRange[i + 1, 59].Value;
            //    WBook[i].WeldProcess = WeldingBook.DataBodyRange[i + 1, 57].Value2;
            //    WBook[i].Welders = WeldersSeparator(WeldingBook.DataBodyRange[i + 1, 55].Value2);
            //    WBook[i].WeldersToBlame = WeldersSeparator(WeldingBook.DataBodyRange[i + 1, 56].Value2);
            //    WBook[i].RTProtNum = WeldingBook.DataBodyRange[i + 1, 20].Value2;
            //    WBook[i].UTProtNum = WeldingBook.DataBodyRange[i + 1, 24].Value2;
            //    WBook[i].RTDate = WeldingBook.DataBodyRange[i + 1, 4].Value2;
            //    WBook[i].UTDate = WeldingBook.DataBodyRange[i + 1, 5].Value2;
            //    WBook[i].Result = WeldingBook.DataBodyRange[i + 1, 54].Value2;
            //}

            for (int i = 1; i < WBarray.GetUpperBound(0); i++)
            {
                WBook.Add(new Weld());
                WBook[i - 1].DrawingNum = WBarray[i, 8];
                WBook[i - 1].ISONum = WBarray[i, 9];
                WBook[i - 1].WeldNumber = WBarray[i, 26];
                WBook[i - 1].EndDate = (DateTime)WBarray[i, 65];
                WBook[i - 1].WeldMaterial = WBarray[i, 11];
                WBook[i - 1].WeldDiam = (double)WBarray[i, 58];
                WBook[i - 1].WeldThick = (double)WBarray[i, 59];
                WBook[i - 1].WeldProcess = WBarray[i, 57];
                WBook[i - 1].Welders = WeldersSeparator((string)WBarray[i, 55]);
                WBook[i - 1].WeldersToBlame = WeldersSeparator((string)WBarray[i, 56]);
                WBook[i - 1].RTProtNum = WBarray[i, 20];
                WBook[i - 1].UTProtNum = WBarray[i, 24];
                WBook[i - 1].RTDate = WBarray[i, 4];
                WBook[i - 1].UTDate = WBarray[i, 5];
                WBook[i - 1].Result = WBarray[i, 54];
                WBook[i - 1].Object = WBarray[i, 64];

            }
            return WBook;
        }
        public static List<Welder> WeldersRates()
        {
            List<Weld> WBook = WeldData();
            List<Welder> WelderBase = new List<Welder>();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "Repair Rate Sharp.xlsb") != null)
            {
                wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\002-repair rates\Repair Rate Sharp.xlsb");
            }
            Worksheet ws = wb.Worksheets["unofficial"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range Weldersrange = ws.get_Range("B8:C8", last);
            var WeldersArray = (object[,])Weldersrange.Value;
            for (int i = 1; i < WeldersArray.GetUpperBound(0) + 1; i++)
            {
                WelderBase.Add(new Welder());
                WelderBase[i - 1].WeldersName = (string)WeldersArray[i, 1];
                WelderBase[i - 1].Stamp = (string)WeldersArray[i, 2];
            }
            foreach (Weld ProdWeld in WBook)
            {
                foreach (Welder WelderUnit in WelderBase)
                {
                    if (ProdWeld.Welders.Contains(WelderUnit.Stamp))
                    {
                        WelderUnit.WelderWelds.Add(ProdWeld);
                    }
                }
            }
            return WelderBase;
        }
        public static void PrintRates()
        {
            RepairForm FormRepair = new RepairForm();
            FormRepair.Show();

            List<Welder> WelderBase = WeldersRates();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            
            

            Workbook wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            Worksheet ws = wb.Worksheets["parameters"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            string ProdObject = ws.Range["B2"].Value;
            DateTime DateMIN = (DateTime)ws.Range["C2"].Value;
            DateTime DateMAX = (DateTime)ws.Range["D2"].Value;
            var ParArray = (object[,])ws.get_Range("B4:B7", last).Value;
            CountRates(WelderBase, DateMIN, DateMAX, ProdObject, ParArray);
            ws = wb.Worksheets["unofficial"];
            for (int i = 0; i < WelderBase.Count-1; i++)
            {
                ws.Cells[i + 8, 4] = WelderBase[i].WelderRates[0].Overall;
                ws.Cells[i + 8, 5] = WelderBase[i].WelderRates[0].NDErates();
                ws.Cells[i + 8, 6] = WelderBase[i].WelderRates[0].Rates();
                ws.Cells[i + 8, 7] = WelderBase[i].WelderRates[0].Accept;
                ws.Cells[i + 8, 8] = WelderBase[i].WelderRates[0].Repair;
                ws.Cells[i + 8, 9] = WelderBase[i].WelderRates[0].Cutout;

                for (int j = 1; j < 6; j++)
                {
                    ws.Cells[i + 8, 10 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1 - 1].Overall + WelderBase[i].WelderRates[j + j * 1].Overall;
                    ws.Cells[i + 8, 11 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1 - 1].Rates();
                    ws.Cells[i + 8, 12 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1 - 1].Accept;
                    ws.Cells[i + 8, 13 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1 - 1].Repair;
                    ws.Cells[i + 8, 14 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1 - 1].Cutout;

                    ws.Cells[i + 8, 15 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1].Rates();
                    ws.Cells[i + 8, 16 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1].Accept;
                    ws.Cells[i + 8, 17 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1].Repair;
                    ws.Cells[i + 8, 18 + j * 9 - 9] = WelderBase[i].WelderRates[j + j * 1].Cutout;
                }
                ws.Cells[i + 8, 55] = WelderBase[i].WelderRates[11].Overall;
                ws.Cells[i + 8, 56] = WelderBase[i].WelderRates[11].Rates();
                ws.Cells[i + 8, 57] = WelderBase[i].WelderRates[11].Accept;
                ws.Cells[i + 8, 58] = WelderBase[i].WelderRates[11].Repair;
                ws.Cells[i + 8, 59] = WelderBase[i].WelderRates[11].Cutout;
                ws.Cells[i + 8, 60] = WelderBase[i].WelderRates[12].Overall;
                ws.Cells[i + 8, 61] = WelderBase[i].WelderRates[12].Rates();
                ws.Cells[i + 8, 62] = WelderBase[i].WelderRates[12].Accept;
                ws.Cells[i + 8, 63] = WelderBase[i].WelderRates[12].Repair;
                ws.Cells[i + 8, 64] = WelderBase[i].WelderRates[12].Cutout;
                //for (int j = 1; j < 13; j++)
                //{
                //}
            }
        }
        public static List<Welder> CountRates(List<Welder> WelderBase, DateTime DateMIN, DateTime DateMAX, string ProdObject, object[,] ParArray)
        {

            foreach (Welder welder in WelderBase)
            {
                foreach (Weld weld in welder.WelderWelds)
                {
                   
                   if (DateMIN <= weld.EndDate && DateMAX >= weld.EndDate && (ProdObject == "All" || ProdObject == (string) weld.Object))
                    {
                        welder.WelderRates[0].Overall = welder.WelderRates[0].Overall + 1;
                        if (weld.RTProtNum != null || weld.RTDate != null || weld.UTProtNum != null || weld.UTDate != null)
                        {
                            switch (weld.Result)
                            {
                                case "accepted":
                                    welder.WelderRates[0].Accept = welder.WelderRates[0].Accept + 1;
                                    break;
                                case "repair":
                                    if (weld.WeldersToBlame.Contains(welder.Stamp))
                                    {
                                        welder.WelderRates[0].Repair = welder.WelderRates[0].Repair + 1;
                                    }                                    
                                    break;
                                case "cutout":
                                    if (weld.WeldersToBlame.Contains(welder.Stamp))
                                    {
                                        welder.WelderRates[0].Cutout = welder.WelderRates[0].Cutout + 1;
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        for (int i = 1; i < ParArray.GetUpperBound(1)+1; i++)
                        {
                            if (((string) weld.WeldMaterial).Contains((string) ParArray[1, i]) && (string) ParArray[2, i] == (string) weld.WeldProcess 
                                && (double)ParArray[3, i] < weld.WeldDiam && (double)ParArray[4, i] >= weld.WeldDiam)
                            {
                                welder.WelderRates[i].Overall = welder.WelderRates[i].Overall + 1;
                                switch (weld.Result)
                                {
                                    case "accepted":
                                        welder.WelderRates[i].Accept = welder.WelderRates[i].Accept + 1;
                                        break;
                                    case "repair":
                                        if (weld.WeldersToBlame.Contains(welder.Stamp))
                                        {
                                            welder.WelderRates[i].Repair = welder.WelderRates[i].Repair + 1;
                                        }
                                        break;
                                    case "cutout":
                                        if (weld.WeldersToBlame.Contains(welder.Stamp))
                                        {
                                            welder.WelderRates[i].Cutout = welder.WelderRates[i].Cutout + 1;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                        }
                    }                
                }
            }
            return WelderBase;
        }

        public static List<string> WeldersSeparator(string WeldersString)
        {
            List<string> UniqWelders;
            UniqWelders = WeldersString.Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries).Distinct().ToList();
            return UniqWelders;
        }
    }   
    public class DataWriter
    {
        public static void WriteData()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            Range headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (int i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            Range dataRange = ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
            xlApp.ActiveChart.SetSourceData(Source: dataRange);
        }

    }
}
