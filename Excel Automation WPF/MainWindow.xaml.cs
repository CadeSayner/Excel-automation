using System.Windows;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace Excel_Automation_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Workbooks workbooks;
        private Workbook wkbk;
        private Worksheet wksht;
        private Worksheet wksht2;

        private string pivotTableName = "PivotTableTest";
        private int lastUsedRow = 0;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                DefaultExt = ".txt",
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                // Start Excel
                InitialiseApplication(openFileDialog.FileName);
                xlApp.DisplayAlerts = false;

                // Delete unused columns
                ClearColumns();

                // Format Columns
                FormatColumns();
                    
                //Deal with emergency cases.
                Emergencies();

                //Instantiate second worksheet
                wksht2 = wkbk.Worksheets.Add();
                wksht2.Name = "Pivot Table";

                // Create Pivot Table
                CreatePivotTable();

                //// Show 
                xlApp.Visible = true;

                //xlApp.DisplayAlerts = true;
                CleanUp();
            }
        }

        private void InitialiseApplication(string fileName)
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            workbooks = xlApp.Workbooks;
            workbooks.OpenText(fileName, XlPlatform.xlWindows, 1, XlTextParsingType.xlFixedWidth);
            wkbk = xlApp.ActiveWorkbook;
            wksht = wkbk.Worksheets[1];
            wksht.Name = "Generated Spreadsheet";

   
            wksht.Copy(wksht);
            Worksheet wksht3 = wkbk.Worksheets[1];
            wksht3.Name = "Original";
        }

        private void ClearColumns()
        {
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["A"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["B"].Delete();
            wksht.Columns["C"].Delete();
            wksht.Columns["D"].Delete();
            wksht.Columns["E"].Delete();
        }

        private void CleanUp()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkbk);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wksht);
        }

        private void FormatColumns()
        {
            // Add columns for split
            InsertColumn("B1", 2);

            wksht.Range["A1"].EntireColumn.ColumnWidth = 30;

            // format text such that commas are placed for delimination
            lastUsedRow = wksht.UsedRange.Rows.Count;

            // Process Date and time columns fields 
            for (int i = 1; i <= lastUsedRow; i++)
            {
                Range r = wksht.Cells[i, 1];
                if (r.Text.Length > 6)
                {
                    string s = r.Text;
                    string date = s.Substring(10, 8);
                    string time = s.Substring(18, 4);
                    string dateTime = date + " " + time;
                    wksht.Cells[i, 1] = dateTime;
                }
            }

            wksht.Range["A1", "A" + wksht.UsedRange.Rows.Count].TextToColumns(wksht.Range["B1"], XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierNone,
                        true, Type.Missing, Type.Missing, false, true, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Remove source column
            wksht.Columns["A"].Delete();

            // Insert adjacent columns
            InsertColumn("D1", 3);

            // Increase column size
            wksht.Range["C1"].EntireColumn.ColumnWidth = 30;

            FormatMno();

            wksht.Range["C1", "C" + wksht.UsedRange.Rows.Count].TextToColumns(wksht.Range["D1"], XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierNone,
            true, Type.Missing, Type.Missing, false, true, Type.Missing, " ", Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            wksht.Columns["C"].Delete();

            InsertRow(1, 1);

            AddTitles();
        }

        private void InsertColumn(string startCol, int amount)
        {
            Range oRng = wksht.Range[startCol];

            for (int i = 0; i < amount; i++)
            {
                oRng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
        }

        private void InsertRow(int rowIndex, int amount)
        {
            Range line = wksht.Rows[rowIndex];
            for (int i = 0; i < amount; i++)
            {
                line.Insert();
            }
        }

        private void AddTitles()
        {
            wksht.Cells[1, 1] = "Date";
            wksht.Cells[1, 2] = "Time";
            wksht.Cells[1, 3] = "Manifesto Number";
            wksht.Cells[1, 4] = "Type";
            wksht.Cells[1, 5] = "Order Number";
            wksht.Cells[1, 6] = "Part Number";
            wksht.Cells[1, 7] = "QTY";
        }

        private void FormatMno()
        {
            // Process Manifesto, type and order number fields 
            for (int i = 1; i <= lastUsedRow; i++)
            {
                Range r = wksht.Cells[i, 3];
                string s = r.Text;
                if (s.Contains("GAXP") != true)
                {
                    string manNumber = s.Substring(0, 6);
                    string modelType = s.Substring(6, 2);
                    string orderNumber_ = s.Substring(8);
                    string orderNumber = orderNumber_.Remove(orderNumber_.Length - 2);
                    string mno = manNumber + " " + modelType + " " + orderNumber;
                    wksht.Cells[i, 3] = mno;
                }

                else
                {
                    string manNumber = s.Substring(0, 6);
                    string modelType = s.Substring(6, 4);
                    string orderNumber_ = s.Substring(10);
                    string orderNumber = orderNumber_.Remove(orderNumber_.Length - 2);
                    string mno = manNumber + " " + modelType + " " + orderNumber;
                    wksht.Cells[i, 3] = mno;
                }
            }
        }

        private void Emergencies()
        {

            string[] dateTimeArr = DateTime.Today.Date.ToShortDateString().Split('/');
            string dateTime = String.Join("", dateTimeArr);
            // Loop through 1st column 
            for(int i = 2; i <= lastUsedRow; i++)
            {
                Range cD = wksht.Cells[i, 1];
                Range cT = wksht.Cells[i, 2];


                if(cD.Text.Length < 8)
                {
                    wksht.Cells[i, 1] = dateTime;
                }

                if(cT.Text == "")
                {
                    wksht.Cells[i, 2] = "N/A (Emergency)";
                }
            }
        }

        private void CreatePivotTable()
        {
            Range oRange = wksht.UsedRange;
            Range oRange2 = wksht2.Range["A1"];

            PivotCache oPivotCache = wkbk.PivotCaches().Add(XlPivotTableSourceType.xlDatabase, oRange);

            PivotTable oPivotTable = (PivotTable)wksht2.PivotTables().Add(oPivotCache, oRange2, "Summary");
            oPivotTable.Name = pivotTableName;
            oPivotTable.RowAxisLayout(XlLayoutRowType.xlTabularRow);

            PivotField oPivotFieldDate = (PivotField)oPivotTable.PivotFields("Date");
            oPivotFieldDate.Orientation = XlPivotFieldOrientation.xlPageField;
            oPivotFieldDate.Name = " Date";

            PivotField oPivotFieldTime = (PivotField)oPivotTable.PivotFields("Time");
            oPivotFieldTime.Orientation = XlPivotFieldOrientation.xlPageField;
            oPivotFieldTime.Name = " Time";

            PivotField oPivotFieldType = (PivotField)oPivotTable.PivotFields("Type");
            oPivotFieldType.Orientation = XlPivotFieldOrientation.xlRowField;
            oPivotFieldType.Name = "Model Type";

            PivotField oPivotFieldManifestoNumber = (PivotField)oPivotTable.PivotFields("Part Number");
            oPivotFieldManifestoNumber.Orientation = XlPivotFieldOrientation.xlRowField;
            oPivotFieldManifestoNumber.Name = "Part Number";

            PivotField oPivotFieldqty = (PivotField)oPivotTable.PivotFields("QTY");
            oPivotFieldqty.Orientation = XlPivotFieldOrientation.xlDataField;
        }
    }
}
