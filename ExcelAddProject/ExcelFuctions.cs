using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddProject
{
    //public static class MyFunctions
    //{
    //    [ExcelFunction(Description = "My first Excel-DNA function")]
    //    public static string MyFirstFunction(string name)
    //    {
    //        return "Hello " + name;
    //    }
    //    [ExcelFunction(Description = "Joins a string to a number", Category = "My functions")]
    //    public static string JoinThem([ExcelArgument(Description = "Input string",Name = "Word",AllowReference =true)] string str, [ExcelArgument(Description = "Input number", AllowReference = true)] double val)
    //    {
    //        return str + val;
    //    }
    //    [ExcelFunction(Description = "Multiplies two numbers", Category = "Useful functions")]
    //    public static double MultiplyThem(double x, double y)
    //    {
    //        return x * y;
    //    }
    //    [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
    //    public static double AddThem(
    //[ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")]
    //double v1,
    //[ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]
    //double v2)
    //    {
    //        return v1 + v2;
    //    }
    //}
    public class RepairRates
    {
        public static List<Weld> WeldData(bool WBcalc, bool Official)
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
            //System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Хотите обновить WB сборку перед расчетом брака", "Repair Rate Counter 9000", System.Windows.Forms.MessageBoxButtons.YesNo);
            //if (dialogResult == System.Windows.Forms.DialogResult.Yes)
            ((Worksheet)wb.Worksheets["All"]).AutoFilter.ShowAllData();
            Range Selection = (((Worksheet)wb.Worksheets["All"]).ListObjects["All"].HeaderRowRange.Find("Resultat"));
            if (WBcalc)
            {
                ((Worksheet)wb.Worksheets["NEWS BASE"]).ListObjects["NEWS_BASE"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Workshop"]).ListObjects["Workshop"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Erection"]).ListObjects["Erection"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["Flare"]).ListObjects["Flare"].QueryTable.Refresh(false);
                ((Worksheet)wb.Worksheets["All"]).ListObjects["All"].QueryTable.Refresh(false);
                string Numberum = Selection.Address[false, false, XlReferenceStyle.xlA1, false];
                if ((((Worksheet)wb.Worksheets["All"]).ListObjects["All"].Range.Rows.Count - Selection.End[XlDirection.xlDown].Row) > 1)
                {
                    Selection = Selection.get_Offset(1, 0);
                    Selection = Selection.get_Resize(1, Selection.Column + Selection.End[XlDirection.xlToRight].Column);
                    Selection.Copy();
                    Selection = Selection.get_Offset(Selection.End[XlDirection.xlDown].Row - Selection.Row, 0);
                    Selection = Selection.get_Resize(Selection.End[XlDirection.xlDown].Row - Selection.Row + 1, Selection.Columns.Count);
                    Selection.Select();
                    ((Worksheet)xlApp.ActiveSheet).Paste();
                    ((Worksheet)xlApp.ActiveSheet).Calculate();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Новых стыков не найдено");
                }
            }
            else if ((((Worksheet)wb.Worksheets["All"]).ListObjects["All"].Range.Rows.Count - Selection.End[XlDirection.xlDown].Row) > 1)
            {
                ((Worksheet)wb.Worksheets["All"]).Activate();
                ((Worksheet)wb.Worksheets["All"]).AutoFilter.ShowAllData();
                Selection = Selection.get_Offset(1, 0);
                Selection = Selection.get_Resize(1, Selection.Column + Selection.End[XlDirection.xlToRight].Column);
                Selection.Copy();
                Selection = Selection.get_Offset(Selection.End[XlDirection.xlDown].Row - Selection.Row, 0);
                Selection = Selection.get_Resize(Selection.End[XlDirection.xlDown].Row - Selection.Row + 1, Selection.Columns.Count);
                Selection.Select();
                ((Worksheet)xlApp.ActiveSheet).Paste();
                ((Worksheet)xlApp.ActiveSheet).Calculate();
            }
            //else if (dialogResult == System.Windows.Forms.DialogResult.No)
            //{

            //}
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
            int Result = 55;
            if (Official) Result = 18;


            for (int i = 1; i < WBarray.GetUpperBound(0); i++)
            {
                WBook.Add(new Weld());
                WBook[i - 1].WeldDiamInch = (double)WBarray[i, 7];
                WBook[i - 1].DrawingNum = Convert.ToString(WBarray[i, 8]);
                WBook[i - 1].ISONum = Convert.ToString(WBarray[i, 9]);
                WBook[i - 1].WeldNumber = Convert.ToString(WBarray[i, 26]);
                WBook[i - 1].EndDate = (DateTime)WBarray[i, 66];
                WBook[i - 1].WeldMaterial = Convert.ToString(WBarray[i, 11]);
                WBook[i - 1].WeldDiam = (double)WBarray[i, 59];
                WBook[i - 1].WeldThick = (double)WBarray[i, 60];
                WBook[i - 1].WeldProcess = Convert.ToString(WBarray[i, 58]);
                WBook[i - 1].Welders = WeldersSeparator((string)WBarray[i, 56]);
                WBook[i - 1].WeldersToBlame = WeldersSeparator((string)WBarray[i, 57]);
                WBook[i - 1].RTProtNum = Convert.ToString(WBarray[i, 20]);
                WBook[i - 1].UTProtNum = Convert.ToString(WBarray[i, 24]);
                WBook[i - 1].RTDate = Convert.ToString(WBarray[i, 4]);
                WBook[i - 1].UTDate = Convert.ToString(WBarray[i, 5]);
                WBook[i - 1].NDEcontrol = (bool)WBarray[i, 70];
                WBook[i - 1].Result = Convert.ToString(WBarray[i, Result]);
                WBook[i - 1].Object = Convert.ToString(WBarray[i, 65]);
                WBook[i - 1].IsRepair = (bool)WBarray[i, 72];
            }
            return WBook;
        }
        public static List<Welder> WelderNameFiller(string ReportType)
        {
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
            Worksheet ws = new Worksheet();
            ws = wb.Worksheets[ReportType];
            string DataStr;
            switch (ReportType)
            {
                case "Simple Rates":
                    {
                        DataStr = "B8:C8";
                    }
                    break;
                case "Repair Rates":
                    {
                        DataStr = "B8:C8";
                    }
                    break;
                case "Disqual Rates":
                    {
                        DataStr = "B8:C8";
                    }
                    break;
                case "Qual summary":
                    {
                        DataStr = "B8:C8";
                    }
                    break;
                default:
                    {
                    }
                    break;
            }
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range Weldersrange = ws.get_Range("B8:C8", last);
            var WeldersArray = (object[,])Weldersrange.Value;
            for (int i = 1; i < WeldersArray.GetUpperBound(0) + 1; i++)
            {
                WelderBase.Add(new Welder());
                WelderBase[i - 1].WeldersName = (string)WeldersArray[i, 1];
                WelderBase[i - 1].Stamp = (string)WeldersArray[i, 2];
            }
            return WelderBase;
        }
        public static List<Welder> WeldersRates(List<Weld> WBook, List<Welder> WelderBase)
        {
            //List<Welder> WelderBase = new List<Welder>();
            //Application xlApp = (Application)ExcelDnaUtil.Application;
            //xlApp.DisplayAlerts = false;
            //Workbook wb;
            //if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "Repair Rate Sharp.xlsb") != null)
            //{
            //    wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            //}
            //else
            //{
            //    wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\002-repair rates\Repair Rate Sharp.xlsb");
            //}
            //Worksheet ws = new Worksheet();
            //ws = wb.Worksheets[ReportType];
            //Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //Range Weldersrange = ws.get_Range("B8:C8", last);
            //var WeldersArray = (object[,])Weldersrange.Value;
            //for (int i = 1; i < WeldersArray.GetUpperBound(0) + 1; i++)
            //{
            //    WelderBase.Add(new Welder());
            //    WelderBase[i - 1].WeldersName = (string)WeldersArray[i, 1];
            //    WelderBase[i - 1].Stamp = (string)WeldersArray[i, 2];
            //}
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
        public static List<Welder> WeldersQuals(List<QualSimp> QualBase, List<Welder> WelderBase)
        {
            foreach (QualSimp Qual in QualBase)
            {
                foreach (Welder welder in WelderBase)
                {
                    if (welder.Stamp != null && Qual.Stamp.Contains((string)welder.Stamp))
                    {
                        welder.QualSimpl.Add(Qual);
                    }
                }
            }
            return WelderBase;
        }
        public static string[,] CountQuals(List<Welder> WelderBase)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            Worksheet ws = wb.Worksheets["parameters"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            var ParArray = (object[,])ws.get_Range("B14:P17").Value;
            string[,] DataString = new string[WelderBase.Count, ParArray.GetLength(1)];
            for (int j = 0; j < WelderBase.Count - 1; j++)
            {
                foreach (QualSimp qual in WelderBase[j].QualSimpl)
                {
                    for (int i = 0; i < ParArray.GetLength(1); i++)
                    {
                        if (((String)ParArray[1, i + 1]).Contains(qual.MaterialGroup) && (String)ParArray[2, i + 1] == qual.KSSProcess && ((String)ParArray[3, i + 1]).Contains(qual.Position) && (double)ParArray[4, i + 1] <= (double)qual.QualDimentions.DiameterMin)
                        {
                            DataString[j, i] = qual.DLname;
                        }
                    }
                }
            }
            return DataString;
        }
        public static void PrintQuals(string[,] DataString)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            ((Worksheet)wb.Worksheets["Qual summary"]).Activate();
            Worksheet ws = wb.Worksheets["Qual summary"];
            Range c1 = ws.Cells[8, 4];
            Range c2 = ws.Cells[DataString.GetLength(0) + 8, DataString.GetLength(1) + 4];
            Range range = xlApp.get_Range(c1, c2);
            range.Value = DataString;
        }
        public static List<Welder> WeldersRatesTL(bool WBcalc, bool Official)
        {
            List<Weld> WBook = WeldData(WBcalc, Official);
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
            Worksheet ws = new Worksheet();
            ws = wb.Worksheets["Timeline"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            Range Weldersrange = ws.get_Range("A6:B6", last);
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
        public static void PrintRates(DateTime DateMIN, DateTime DateMAX, string ProdObject, bool WBcalc, bool Official, string ReportType, bool NoRepairs, bool Volume)
        {
            bool NoDiameter;
            string DataSource;
            List<Welder> WelderBase = new List<Welder>();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            Worksheet ws = wb.Worksheets["parameters"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //string ProdObject = ws.Range["B2"].Value;
            //DateTime DateMIN = (DateTime)ws.Range["C2"].Value;
            //DateTime DateMAX = (DateTime)ws.Range["D2"].Value;
            switch (ReportType)
            {
                case "Simple Rates":
                    {
                        NoDiameter = true;
                        WelderBase = WeldersRates(WeldData(WBcalc, Official), WelderNameFiller(ReportType));
                        CountRates(WelderBase, DateMIN, DateMAX, ProdObject, (object[,])ws.get_Range("B10:H11").Value, NoDiameter, NoRepairs, Volume);
                        ((Worksheet)wb.Worksheets["Simple Rates"]).Activate();
                        for (int i = 0; i < WelderBase.Count - 1; i++)
                        {
                            for (int j = 0; j < 8; j++)
                            {
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 10 + j * 6 - 6] = WelderBase[i].WelderRates[j].Overall;
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 11 + j * 6 - 6] = WelderBase[i].WelderRates[j].NDErates();
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 12 + j * 6 - 6] = WelderBase[i].WelderRates[j].Rates();
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 13 + j * 6 - 6] = WelderBase[i].WelderRates[j].Accept;
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 14 + j * 6 - 6] = WelderBase[i].WelderRates[j].Repair;
                                ((Worksheet)wb.Worksheets["Simple Rates"]).Cells[i + 8, 15 + j * 6 - 6] = WelderBase[i].WelderRates[j].Cutout;
                            }
                        }
                    }
                    break;
                case "Repair Rates":
                    {
                        NoDiameter = false;
                        WelderBase = WeldersRates(WeldData(WBcalc, Official), WelderNameFiller(ReportType));
                        CountRates(WelderBase, DateMIN, DateMAX, ProdObject, (object[,])ws.get_Range("B4:B7", last).Value, NoDiameter, NoRepairs, Volume);
                        ((Worksheet)wb.Worksheets["Repair Rates"]).Activate();
                        for (int i = 0; i < WelderBase.Count - 1; i++)
                        {
                            ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 4] = WelderBase[i].WelderRates[0].Overall;
                            ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 5] = WelderBase[i].WelderRates[0].NDErates();
                            ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 6] = WelderBase[i].WelderRates[0].Rates();
                            ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 7] = WelderBase[i].WelderRates[0].Accept;
                            ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 8] = WelderBase[i].WelderRates[0].Repair + WelderBase[i].WelderRates[0].Cutout;

                            for (int j = 1; j < 16; j++)
                            {
                                ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 9 + j * 3 - 3] = WelderBase[i].WelderRates[j].Rates();
                                ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 10 + j * 3 - 3] = WelderBase[i].WelderRates[j].Accept;
                                ((Worksheet)wb.Worksheets["Repair Rates"]).Cells[i + 8, 11 + j * 3 - 3] = WelderBase[i].WelderRates[j].Repair + WelderBase[i].WelderRates[j].Cutout;
                            }
                        }
                    }
                    break;
                case "Disqual Rates":
                    {
                        NoDiameter = false;
                        WelderBase = WeldersRates(WeldData(WBcalc, Official), WelderNameFiller(ReportType));
                        CountRates(WelderBase, DateMIN, DateMAX, ProdObject, (object[,])ws.get_Range("B4:B7", last).Value, NoDiameter, NoRepairs, Volume);
                        ((Worksheet)wb.Worksheets["Disqual Rates"]).Activate();
                        for (int i = 0; i < WelderBase.Count - 1; i++)
                        {
                            ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 4] = WelderBase[i].WelderRates[0].Overall;
                            ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 5] = WelderBase[i].WelderRates[0].NDErates();
                            ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 6] = WelderBase[i].WelderRates[0].Rates();
                            ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 7] = WelderBase[i].WelderRates[0].NDETotal();

                            for (int j = 1; j < 16; j++)
                            {
                                ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 8 + j * 2 - 2] = WelderBase[i].WelderRates[j].Rates();
                                ((Worksheet)wb.Worksheets["Disqual Rates"]).Cells[i + 8, 9 + j * 2 - 2] = WelderBase[i].WelderRates[j].NDETotal();
                            }
                        }
                    }
                    break;
                default:
                    {
                    }
                    break;
            }
            xlApp.DisplayAlerts = true;
        }
        public static void PrintratesCOK(Dictionary<string, RatesContainerCOK> RatesContainerCOKAll, Dictionary<string, RatesContainerCOK> RatesContainerCOKMonth)
        {
            int i = 7;
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
            ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Activate();
            foreach (KeyValuePair<string, RatesContainerCOK> RatesContainer in RatesContainerCOKMonth)
            {
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 5] = RatesContainer.Value.Overall;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 6] = RatesContainer.Value.NDEOverall;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 8] = RatesContainer.Value.NDEDone;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 10] = RatesContainer.Value.NDEReject;
                i = i + 1;
            }
            i = i + 3;
            foreach (KeyValuePair<string, RatesContainerCOK> RatesContainer in RatesContainerCOKAll)
            {
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 5] = RatesContainer.Value.Overall;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 6] = RatesContainer.Value.NDEOverall;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 8] = RatesContainer.Value.NDEDone;
                ((Worksheet)wb.Worksheets["ПБР Тенгиз"]).Cells[i, 10] = RatesContainer.Value.NDEReject;
                i = i + 1;
            }
            xlApp.DisplayAlerts = true;
        }
        public static Dictionary<string, RatesContainerCOK> CountRatesCOK(DateTime StartDate, DateTime EndDate, List<Weld> WBook, bool InchCount)
        {
            Dictionary<string, RatesContainerCOK> RatesContainerCOKAll = new Dictionary<string, RatesContainerCOK>();
            List<string> Materials = new List<string> { "LTCS", "SS", "F22", "ALLOY" };
            List<string> Objects = new List<string> { "Workshop", "Erection" };
            double inchcounter = 1;

            foreach (string Material in Materials)
            {
                foreach (string Object in Objects)
                {
                    RatesContainerCOKAll.Add(Material + Object, new RatesContainerCOK());
                }
            }
            foreach (Weld weld in WBook)
            {
                if (InchCount)
                {
                    inchcounter = weld.WeldDiamInch;
                }
                foreach (string Material in Materials)
                {
                    foreach (string Object in Objects)
                    {
                        if (weld.EndDate <= EndDate && weld.EndDate >= StartDate && weld.IsRepair == false && weld.WeldMaterial.ToString().Contains(Material) && weld.Object.ToString() == Object)
                        {
                            RatesContainerCOKAll[Material + Object].Overall = RatesContainerCOKAll[Material + Object].Overall + 1 * inchcounter;
                            if (weld.NDEcontrol)
                            {
                                RatesContainerCOKAll[Material + Object].NDEOverall = RatesContainerCOKAll[Material + Object].NDEOverall + 1 * inchcounter;
                                if (weld.Result.ToString() != "0")
                                {
                                    RatesContainerCOKAll[Material + Object].NDEDone = RatesContainerCOKAll[Material + Object].NDEDone + 1 * inchcounter;
                                    if (weld.Result.ToString() == "cutout" || weld.Result.ToString() == "repair")
                                    {
                                        RatesContainerCOKAll[Material + Object].NDEReject = RatesContainerCOKAll[Material + Object].NDEReject + 1 * inchcounter;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return RatesContainerCOKAll;
        }
        public static void PrintFirstThree()
        {
            List<Welder> WelderBase = WeldersRates(WeldData(false, true), WelderNameFiller("Repair Rates"));
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            ((Worksheet)wb.Worksheets["First Three"]).Activate();
            foreach (Welder welder in WelderBase)
            {
                welder.WelderWelds.Sort((x, y) => DateTime.Compare(x.EndDate, y.EndDate));
            }
            for (int i = 0; i < WelderBase.Count - 1; i++)
            {
                ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 1] = i + 1;
                ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 2] = WelderBase[i].WeldersName;
                ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 3] = WelderBase[i].Stamp;
                for (int j = 0; j < WelderBase[i].WelderWelds.Count - 1 & j <= 9; j++)
                {
                    ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 7 + j * 3 - 3] = WelderBase[i].WelderWelds[j].Result;
                    ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 8 + j * 3 - 3] = WelderBase[i].WelderWelds[j].RTProtNum;
                    ((Worksheet)wb.Worksheets["First Three"]).Cells[i + 8, 9 + j * 3 - 3] = WelderBase[i].WelderWelds[j].EndDate;
                }
            }
            xlApp.DisplayAlerts = true;
        }
        public static List<Welder> CountRates(List<Welder> WelderBase, DateTime DateMIN, DateTime DateMAX, string ProdObject, object[,] ParArray, bool NoDiameter, bool NoRepair, bool volume = false)
        {
            bool Repair, Diameter = false;
            double VolumeCount, VolumeCountR;
            int counter;
            foreach (Welder welder in WelderBase)
            {
                counter = 0;
                foreach (Weld weld in welder.WelderWelds)
                {

                    if (volume) { VolumeCount = weld.WeldDiam * weld.WeldThick / 1000; VolumeCountR = VolumeCount * 0.3; }
                    else { VolumeCount = 1; VolumeCountR = 1; }
                    if (NoRepair) Repair = false;
                    else Repair = weld.IsRepair;
                    if (NoDiameter) Diameter = true;
                    if (DateMIN <= weld.EndDate && DateMAX >= weld.EndDate && (ProdObject == "All" || ProdObject == (string)weld.Object) && Repair != true)
                    {
                        welder.WelderRates[0].Overall = welder.WelderRates[0].Overall + 1 * VolumeCount;
                        if (weld.NDEcontrol)
                        {
                            switch (weld.Result)
                            {
                                case "accepted":
                                    welder.WelderRates[0].Accept = welder.WelderRates[0].Accept + 1 * VolumeCount;
                                    break;
                                case "repair":
                                    if (weld.WeldersToBlame.Contains(welder.Stamp))
                                    {
                                        welder.WelderRates[0].Repair = welder.WelderRates[0].Repair + 1 * VolumeCountR;
                                    }
                                    break;
                                case "cutout":
                                    if (weld.WeldersToBlame.Contains(welder.Stamp))
                                    {
                                        welder.WelderRates[0].Cutout = welder.WelderRates[0].Cutout + 1 * VolumeCount;
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        for (int i = 1; i < ParArray.GetUpperBound(1) + 1; i++)
                        {
                            if (((string)weld.WeldMaterial).Contains((string)ParArray[1, i]) && (string)ParArray[2, i] == (string)weld.WeldProcess &&
                                (Diameter || ((double)ParArray[3, i] < weld.WeldDiam && (double)ParArray[4, i] >= weld.WeldDiam)) && Repair != true)
                            {
                                welder.WelderRates[i].Overall = welder.WelderRates[i].Overall + 1 * VolumeCount;
                                if (weld.NDEcontrol)
                                {
                                    switch (weld.Result)
                                    {
                                        case "accepted":
                                            welder.WelderRates[i].Accept = welder.WelderRates[i].Accept + 1 * VolumeCount;
                                            break;
                                        case "repair":
                                            if (weld.WeldersToBlame.Contains(welder.Stamp))
                                            {
                                                welder.WelderRates[i].Repair = welder.WelderRates[i].Repair + 1 * VolumeCountR;
                                            }
                                            break;
                                        case "cutout":
                                            if (weld.WeldersToBlame.Contains(welder.Stamp))
                                            {
                                                welder.WelderRates[i].Cutout = welder.WelderRates[i].Cutout + 1 * VolumeCount;
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
            }
            return WelderBase;
        }
        public static string MaterialConvert(string Material)
        {
            switch (Material)
            {
                case "P.№1":
                    Material = "LTCS";
                    break;
                case "P.№8":
                    Material = "SS";
                    break;
                case "P.№5A":
                    Material = "F22";
                    break;
                case "43":
                    Material = "ALLOY";
                    break;
                case "45":
                    Material = "ALLOY";
                    break;
                default:
                    Material = "";
                    break;
            }
            return Material;
        }
        public static Welder CountRatesTimeline(string Material, Welder Welder, List<KSS> QualKSS, List<KSS> TestKSS, DateTime DateMIN, DateTime DateMAX)
        {
            DateTime MinimumDate = DateMIN;
            foreach (KSS TempKSS in QualKSS)
            {
                foreach (TimelineRates timelineRates in Welder.TimelineRates)
                {
                    if (DateMIN <= TempKSS.EndDate && DateMAX >= TempKSS.EndDate && ((string)TempKSS.MaterialPNo1).Contains(MaterialConvert(Material)))
                    {
                        switch (TempKSS.Dlstatus)
                        {
                            case "передопуск":
                                timelineRates.isGapQual = true;
                                break;
                            case "подписан":
                                timelineRates.isQual = true;
                                break;
                        }
                        break;
                    }
                }
            }
            foreach (KSS TempKSS in TestKSS)
            {
                foreach (TimelineRates timelineRates in Welder.TimelineRates)
                {
                    if (DateMIN <= TempKSS.EndDate && DateMAX >= TempKSS.EndDate && ((string)TempKSS.MaterialPNo1).Contains(MaterialConvert(Material)))
                    {
                        timelineRates.isRequal = true;
                        break;
                    }
                }
            }
            foreach (Weld weld in Welder.WelderWelds)
            {
                foreach (TimelineRates timelineRates in Welder.TimelineRates)
                {

                    if (MinimumDate <= weld.EndDate && DateMAX >= weld.EndDate && ((string)weld.WeldMaterial).Contains(Material))
                    {
                        timelineRates.RatesContainer.Overall = timelineRates.RatesContainer.Overall + 1;
                        if (weld.NDEcontrol)
                        {
                            switch (weld.Result)
                            {
                                case "accepted":
                                    timelineRates.RatesContainer.Accept = timelineRates.RatesContainer.Accept + 1;
                                    break;
                                case "repair":
                                    if (weld.WeldersToBlame.Contains(Welder.Stamp))
                                    {
                                        timelineRates.RatesContainer.Repair = timelineRates.RatesContainer.Repair + 1;
                                    }
                                    break;
                                case "cutout":
                                    if (weld.WeldersToBlame.Contains(Welder.Stamp))
                                    {
                                        timelineRates.RatesContainer.Cutout = timelineRates.RatesContainer.Cutout + 1;
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    if (timelineRates.isRequal)
                    {
                        MinimumDate = timelineRates.StartDate;
                    }
                }
            }
            return Welder;
        }
        public static List<Welder> TimelineCount(DateTime DateMIN, DateTime DateMAX, string Material, bool WBcalc, bool Official)
        {
            List<KSS> QualKSS = WeldersQualifications.GetKSSes(false);
            List<KSS> TestKSS = WeldersQualifications.GetKSSes(true);
            List<Welder> WelderBase = WeldersRatesTL(WBcalc, Official);
            DateTime TempDate;
            int count = 0;
            WelderBase = WelderKSS(WelderBase, TestKSS, true);
            WelderBase = WelderKSS(WelderBase, QualKSS, false);
            for (int i = -1; DateMIN.AddDays(i) < DateMAX; i = i + 7)
            {
                foreach (Welder welder in WelderBase)
                {
                    welder.TimelineRates.Add(new TimelineRates());
                    welder.TimelineRates[count].StartDate = DateMIN.AddDays(i);
                    welder.TimelineRates[count].EndDate = DateMIN.AddDays(i).AddDays(7);
                }
                count = count + 1;
            }
            for (int i = 0; i < WelderBase.Count; i++)
            {
                WelderBase[i] = CountRatesTimeline(Material, WelderBase[i], QualKSS, TestKSS, DateMIN, DateMAX);
            }
            return WelderBase;
        }
        public static void PrintTimeline(List<Welder> WelderBase)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            ((Worksheet)wb.Worksheets["Timeline"]).Activate();
            for (int i = 0; i <= WelderBase.Count - 1; i++)
            {
                for (int j = 0; j < WelderBase[i].TimelineRates.Count - 1; j++)
                {
                    ((Worksheet)wb.Worksheets["Timeline"]).Cells[i + 6, 2 + j] = WelderBase[i].TimelineRates[j].RatesContainer.Rates();
                    if (WelderBase[i].TimelineRates[j].isRequal) ((Range)((Worksheet)wb.Worksheets["Timeline"]).Cells[i + 6, 2 + j]).Interior.Color = XlRgbColor.rgbAqua;
                }
            }
            xlApp.DisplayAlerts = true;
        }
        public static List<Welder> WelderKSS(List<Welder> WelderBase, List<KSS> InputKSS, bool istest)
        {
            foreach (KSS weld in InputKSS)
            {
                foreach (Welder WelderUnit in WelderBase)
                {
                    if (weld.Stamp.Contains((string)WelderUnit.Stamp))
                    {
                        if (istest) WelderUnit.TestKSS.Add(weld);
                        else WelderUnit.QualKSS.Add(weld);
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
    public class WeldersQualifications
    {
        public static List<QualSimp> GetQuals()
        {
            List<QualSimp> QualBase = new List<QualSimp>();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "FM-401.06.xlsb") != null)
            {
                wb = xlApp.Workbooks["FM-401.06.xlsb"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\004-qualifications\FM-401.06.xlsb", Password: "123", ReadOnly: true);
            }
            Worksheet ws = wb.Worksheets["QualList"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            var QualArray = (object[,])(ws.get_Range("C11:AU11", last)).Value;
            int count = 0;
            for (int i = 1; i < QualArray.GetUpperBound(0) - 1; i++)
            {
                if ((string)QualArray[i, 45] == "Допуск активен")
                {
                    QualBase.Add(new QualSimp());
                    QualBase[count].WelderName = Convert.ToString(QualArray[i, 3]);
                    QualBase[count].Stamp = Convert.ToString(QualArray[i, 4]);
                    QualBase[count].DLname = Convert.ToString(QualArray[i, 7]);
                    QualBase[count].Position = Convert.ToString(QualArray[i, 24]);
                    QualBase[count].Dlstart = (DateTime)QualArray[i, 8];
                    QualBase[count].KSSProcess = Convert.ToString(QualArray[i, 9]);
                    QualBase[count].WelderProcess = Convert.ToString(QualArray[i, 10]);
                    if (Convert.ToString(QualArray[i, 31]) == "") QualArray[i, 31] = 0;
                    if (Convert.ToString(QualArray[i, 16]) == "") QualArray[i, 16] = 0;
                    QualBase[count].QualDimentions.DiameterMin = Convert.ToDouble(QualArray[i, 31]);
                    QualBase[count].QualDimentions.ThiknessMin = Convert.ToDouble(QualArray[i, 16]);
                    if (Convert.ToString(QualArray[i, 19]) == "unl") QualArray[i, 19] = 999;
                    if (Convert.ToString(QualArray[i, 19]) == "") QualArray[i, 19] = 0;
                    if (Convert.ToString(QualArray[i, 17]) == "") QualArray[i, 17] = 0;
                    QualBase[count].QualDimentions.ThiknessMax = Convert.ToDouble(QualArray[i, 17]) + Convert.ToDouble(QualArray[i, 19]);
                    QualBase[count].MaterialGroup = Convert.ToString(QualArray[i, 41]);
                    count++;
                }
            }
            return QualBase;
        }
        public static List<KSS> GetKSSes(bool istest)
        {
            List<KSS> KSSBase = new List<KSS>();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "FM-401.06 Sharp.xlsb") != null)
            {
                wb = xlApp.Workbooks["FM-401.06 Sharp.xlsb"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\004-qualifications\FM-401.06 Sharp.xlsb", Password: "123");
            }
            Worksheet ws = new Worksheet();
            if (istest) ws = wb.Worksheets["QWPJ"];
            else ws = wb.Worksheets["TEST"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            var KSSArray = (object[,])(ws.get_Range("C11:BC11", last)).Value;
            for (int i = 1; i < KSSArray.GetUpperBound(0) - 1; i++)
            {
                KSSBase.Add(new KSS());
                KSSBase[i - 1].WelderName = Convert.ToString(KSSArray[i, 1]);
                KSSBase[i - 1].Stamp = Convert.ToString(KSSArray[i, 4]);
                KSSBase[i - 1].KSSNumber = Convert.ToString(KSSArray[i, 2]);
                KSSBase[i - 1].Diametermm = (double)KSSArray[i, 15];
                KSSBase[i - 1].Diameterinch = Convert.ToString(KSSArray[i, 14]);
                KSSBase[i - 1].Thickness.EnterThickness = Convert.ToString(KSSArray[i, 16]);
                KSSBase[i - 1].Material1 = Convert.ToString(KSSArray[i, 7]);
                KSSBase[i - 1].MaterialHeat1 = Convert.ToString(KSSArray[i, 10]);
                KSSBase[i - 1].MaterialPNo1 = Convert.ToString(KSSArray[i, 8]);
                KSSBase[i - 1].MaterialGroup1 = Convert.ToString(KSSArray[i, 9]);
                KSSBase[i - 1].Material2 = Convert.ToString(KSSArray[i, 11]);
                KSSBase[i - 1].MaterialHeat2 = Convert.ToString(KSSArray[i, 14]);
                KSSBase[i - 1].MaterialPNo2 = Convert.ToString(KSSArray[i, 12]);
                KSSBase[i - 1].MaterialGroup2 = Convert.ToString(KSSArray[i, 13]);
                KSSBase[i - 1].Position = Convert.ToString(KSSArray[i, 23]);
                KSSBase[i - 1].KSSProcess = Convert.ToString(KSSArray[i, 18]);
                KSSBase[i - 1].WelderProcess = Convert.ToString(KSSArray[i, 19]);
                KSSBase[i - 1].WeldLayers = Convert.ToString(KSSArray[i, 20]);
                KSSBase[i - 1].Prosedure = Convert.ToString(KSSArray[i, 37]);
                KSSBase[i - 1].WeldType = Convert.ToString(KSSArray[i, 22]);
                KSSBase[i - 1].WPS = Convert.ToString(KSSArray[i, 5]);
                KSSBase[i - 1].EndDate = (DateTime)KSSArray[i, 24];
                //KSSBase[i - 1].VTrequestNumber = Convert.ToString(KSSArray[i, 25]);
                //KSSBase[i - 1].VTrequestDate = (DateTime)KSSArray[i, 26];
                //KSSBase[i - 1].VTprotocolNumber = Convert.ToString(KSSArray[i, 27]);
                //KSSBase[i - 1].VTprotocolDate = (DateTime)KSSArray[i, 28];
                //KSSBase[i - 1].NDTrequestNumber = Convert.ToString(KSSArray[i, 29]);
                //KSSBase[i - 1].NDTrequestDate = (DateTime)KSSArray[i, 30];
                //KSSBase[i - 1].NDTprotocolNumber = Convert.ToString(KSSArray[i, 31]);
                //KSSBase[i - 1].NDTprotocolDate = (DateTime)KSSArray[i, 32];
                //KSSBase[i - 1].NDTType = Convert.ToString(KSSArray[i, 33]);
                //KSSBase[i - 1].MECHprotocolNumber = Convert.ToString(KSSArray[i, 46]);
                //KSSBase[i - 1].MECHprotocolDate = (DateTime)KSSArray[i, 47];
                //KSSBase[i - 1].Tensile = Convert.ToString(KSSArray[i, 38]);
                //KSSBase[i - 1].TensileWM = Convert.ToString(KSSArray[i, 39]);
                //KSSBase[i - 1].Bend = Convert.ToString(KSSArray[i, 40]);
                //KSSBase[i - 1].Impactresult = Convert.ToString(KSSArray[i, 41]);
                //KSSBase[i - 1].Macro = Convert.ToString(KSSArray[i, 42]);
                //KSSBase[i - 1].Hardnessresult = Convert.ToString(KSSArray[i, 43]);
                KSSBase[i - 1].Dlstatus = Convert.ToString(KSSArray[i, 32]);
                KSSBase[i - 1].DLname = Convert.ToString(KSSArray[i, 33]);
                KSSBase[i - 1].Dldate = (DateTime)KSSArray[i, 34];
            }
            return KSSBase;
        }
        ////public static List<Qualification> GetQualification()
        ////{ }
        public static List<WPS> GetWPS()
        {
            List<WPS> WPSBase = new List<WPS>();
            Application xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.DisplayAlerts = false;
            Workbook wb;
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "WPS LTCS, SS, ALLOY, F22.xlsx") != null)
            {
                wb = xlApp.Workbooks["Repair Rate Sharp.xlsb"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(@"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\009-WPQr_WPS\002-WPQr_WPS(р) ПБР Тенгиз 3GI\2. WPS\WPS LTCS, SS, ALLOY, F22.xlsx");
            }
            Worksheet ws = wb.Worksheets["List PQR"];
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            var KSSArray = (object[,])(ws.get_Range("C5:BC8", last)).Value;
            return WPSBase;
        }
        public static void WeldQual()
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
            ((Worksheet)wb.Worksheets["All"]).AutoFilter.ShowAllData();
            Range Selection = (((Worksheet)wb.Worksheets["All"]).ListObjects["All"].HeaderRowRange.Find("Resultat"));
            ((Worksheet)wb.Worksheets["Workshop"]).ListObjects["Workshop"].QueryTable.Refresh(false);
            ((Worksheet)wb.Worksheets["Erection"]).ListObjects["Erection"].QueryTable.Refresh(false);
            ((Worksheet)wb.Worksheets["Flare"]).ListObjects["Flare"].QueryTable.Refresh(false);
            ((Worksheet)wb.Worksheets["All"]).ListObjects["All"].QueryTable.Refresh(false);
            string Numberum = Selection.Address[false, false, XlReferenceStyle.xlA1, false];
            if ((((Worksheet)wb.Worksheets["All"]).ListObjects["All"].Range.Rows.Count - Selection.End[XlDirection.xlDown].Row) > 1)
            {
                Selection = Selection.get_Offset(1, 0);
                Selection = Selection.get_Resize(1, Selection.Column + Selection.End[XlDirection.xlToRight].Column);
                Selection.Copy();
                Selection = Selection.get_Offset(Selection.End[XlDirection.xlDown].Row - Selection.Row, 0);
                Selection = Selection.get_Resize(Selection.End[XlDirection.xlDown].Row - Selection.Row + 1, Selection.Columns.Count);
                Selection.Select();
                ((Worksheet)xlApp.ActiveSheet).Paste();
                ((Worksheet)xlApp.ActiveSheet).Calculate();
            }
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "FM-401.06.xlsb") != null)
            {
                wb = xlApp.Workbooks["FM-401.06.xlsb"];
            }
            else
            {
                wb = xlApp.Workbooks.Open(Filename : @"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\004-qualifications\FM-401.06.xlsb", ReadOnly : true);
            }
            Worksheet ws = wb.Worksheets["QualList"];
            ws.Calculate();
            ws.Copy();
            wb = xlApp.ActiveWorkbook;
            ws = wb.Worksheets["QualList"];
            ws.ShowAllData();
            ws.Columns.EntireColumn.Hidden = false;
            Range wr = ws.get_Range("a1").EntireRow.EntireColumn;
            wr.Copy();
            ws.get_Range("a1").Select();
            xlApp.Selection.PasteSpecial(XlPasteType.xlPasteValues);
            ws.Columns["AV:BE"].Delete();
            ws.Columns["AR"].Delete();
            ws.Columns["AD:AP"].Delete();
            ws.Columns["H"].Delete();
            ws.Columns["G"].Delete();
            ws.Columns["A"].Delete();
            DateTime dt = DateTime.Now;
            xlApp.Workbooks["FM-401.06.xlsb"].Close();
            if (((Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")).Workbooks.Cast<Workbook>().FirstOrDefault(x => x.Name == "Wb сборка.xlsx") != null)
            {
                xlApp.Workbooks["Wb сборка.xlsx"].Close();
            }
            wb.SaveAs(Filename: @"\\veles-srv46-fs\Велесстрой\Служба сварочно-монтажных работ\ОГС\004-qualifications\отчеты\QualList\QualList " + dt.ToShortDateString() + ".xlsx");
            wb.Close();           
            xlApp.DisplayAlerts = true;
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(xlApp);
           
        }
    }

    //public class DataWriter
    //{
    //    public static void WriteData()
    //    {
    //        Application xlApp = (Application)ExcelDnaUtil.Application;

    //        Workbook wb = xlApp.ActiveWorkbook;
    //        if (wb == null)
    //            return;

    //        Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
    //        ws.Range["A1"].Value = "Date";
    //        ws.Range["B1"].Value = "Value";

    //        Range headerRow = ws.Range["A1", "B1"];
    //        headerRow.Font.Size = 12;
    //        headerRow.Font.Bold = true;

    //        // Generally it's faster to write an array to a range
    //        var values = new object[100, 2];
    //        var startDate = new DateTime(2007, 1, 1);
    //        var rand = new Random();
    //        for (int i = 0; i < 100; i++)
    //        {
    //            values[i, 0] = startDate.AddDays(i);
    //            values[i, 1] = rand.NextDouble();
    //        }

    //        ws.Range["A2"].Resize[100, 2].Value = values;
    //        ws.Columns["A:A"].EntireColumn.AutoFit();

    //        // Add a chart
    //        Range dataRange = ws.Range["A1:B101"];
    //        dataRange.Select();
    //        ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
    //        xlApp.ActiveChart.SetSourceData(Source: dataRange);
    //    }
    //}
}
