using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace FillEmptyCells
{
    public partial class Form1 : Form
    {
        Excel._Application app = null;
        Excel._Workbook workbook = null;
        Excel._Worksheet worksheet = null;
        Excel.Range range = null;
        Excel.Range last = null;

        int nRows = 0;
        int nCols = 0;
        int startRowReadExcel = 4;
        int startColumnReadExcel = 10;
        object[,] values;
        object[,] fails;

        public Form1()
        {
            InitializeComponent();
            fileNameLabel.Text = "";
        }


        public void OpenExcel(OpenFileDialog openFileDialog1)
        {
            try
            {
                app = new Excel.Application();
                workbook = app.Workbooks.Add(openFileDialog1.FileName);
                app.Visible = true;
                app.DisplayAlerts = false;
                

                string currentSheet = "Unique Data";
                var excelSheets = workbook.Worksheets;
                worksheet = excelSheets.get_Item(currentSheet);
                worksheet.Activate();


                last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                range = worksheet.get_Range("A" + startRowReadExcel, last);

                values = (object[,])range.Value2;

                nRows = worksheet.UsedRange.Rows.Count;
                nCols = worksheet.UsedRange.Columns.Count;

                if (values[1, 9].ToString() == "Teststeps")
                {
                    CloseExcel();
                    values = null;
                }
            }
            catch
            {
                CloseExcel();

                MessageBox.Show("Eroare", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }



        private void fillEmptyCellsButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "xlsx Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var name = openFileDialog1.FileName.Split('\\');
                fileNameLabel.Text = name[name.Length - 1];
                fileNameLabel.ForeColor = System.Drawing.Color.Black;
                fillEmptyCellsButton.Enabled = false;
                failsMoveAndCountButton.Enabled = false;

                OpenExcel(openFileDialog1);
                if (values != null)
                {
                    ChangeValues();

                    var rangeColumnI = worksheet.get_Range("I6", "I" + nRows);
                    rangeColumnI.NumberFormat = "@";

                    range.set_Value(Missing.Value, values);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;


                    fileNameLabel.ForeColor = System.Drawing.Color.Green;

                    app.ActiveWorkbook.SaveAs(openFileDialog1.FileName);

                    CloseExcel();

                    MessageBox.Show("Modificarea realizata cu succes!", "Info",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                }
                else
                {
                    MessageBox.Show("Excel gresit", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
                fillEmptyCellsButton.Enabled = true;
                failsMoveAndCountButton.Enabled = true;

                values = null;
            }
        }

        

        private void failsMoveAndCountButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "xlsx Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var name = openFileDialog1.FileName.Split('\\');
                fileNameLabel.Text = name[name.Length - 1];
                fileNameLabel.ForeColor = System.Drawing.Color.Black;
                fillEmptyCellsButton.Enabled = false;
                failsMoveAndCountButton.Enabled = false;
                OpenExcel(openFileDialog1);

                if (values != null)
                {
                    int nrOfTotalTests = nRows - startRowReadExcel + 1;


                    //delete duplicates 
                    DeleteDuplicates(values);
                    nRows = worksheet.UsedRange.Rows.Count;
                    int nrOfTotalTestsWODuplicated = nRows - startRowReadExcel + 1;

                    GetAndMoveFails();

                    
                    nRows = worksheet.UsedRange.Rows.Count;
                    int nrOfTotalFails = nRows - startRowReadExcel + 1;


                    DeleteDuplicates(fails);
                    //delete duplicates fails
                    last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    range = worksheet.get_Range("A" + startRowReadExcel, last);
                    fails = (object[,])range.Value;


                    nRows = worksheet.UsedRange.Rows.Count;
                    int nrOfTotalFailsWODuplicated = nRows - startRowReadExcel + 1;

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    range = null;


                    GetAndMoveRetested9();

                    string currentSheet = "Retested units";
                    var excelSheets = workbook.Worksheets;
                    worksheet = excelSheets.get_Item(currentSheet);
                    worksheet.Activate();


                    int startIndexResults = 27;
                    int GColumnNo = 7;
                    worksheet.Cells[startIndexResults++, GColumnNo] = nrOfTotalTests;
                    worksheet.Cells[startIndexResults++, GColumnNo] = nrOfTotalTestsWODuplicated;
                    worksheet.Cells[startIndexResults++, GColumnNo] = nrOfTotalFails;
                    worksheet.Cells[startIndexResults++, GColumnNo] = nrOfTotalFailsWODuplicated;


                    currentSheet = "Unique Data";
                    excelSheets = workbook.Worksheets;
                    worksheet = excelSheets.get_Item(currentSheet);
                    worksheet.Activate();

                    last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    range = worksheet.get_Range("A" + startRowReadExcel, last);

                    //unsort 1 table
                    range.Sort(range.Columns[5], Excel.XlSortOrder.xlAscending);

                    range = null;

                    fileNameLabel.ForeColor = System.Drawing.Color.Green;
                    app.ActiveWorkbook.SaveAs(openFileDialog1.FileName);
                    CloseExcel();


                    MessageBox.Show("Modificarea realizata cu succes!", "Info",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            fillEmptyCellsButton.Enabled = true;
            failsMoveAndCountButton.Enabled = true;

            values = null;
        }

        private void GetAndMoveRetested9()
        {
            string currentSheet = "Unique Data";
            var excelSheets = workbook.Worksheets;
            worksheet = excelSheets.get_Item(currentSheet);
            worksheet.Activate();


            //get all fails
            values = null;

            last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            range = worksheet.get_Range("A1", last);


            //var index = range.Find("start_wt", LookAt: Excel.XlLookAt.xlPart).get_Address(Excel.XlReferenceStyle.xlA1);
            var index = range.Find("start_wt", LookAt: Excel.XlLookAt.xlPart).Column;


            range = worksheet.get_Range("A" + startRowReadExcel, last);
            //sort to have fails on top of table
            //range.Sort(range.Columns[index], Excel.XlSortOrder.xlAscending);
            //range.AutoFilter(1, "-9999999", Excel.XlAutoFilterOperator.xlFilterValues);

            range.AdvancedFilter(Excel.XlFilterAction.xlFilterInPlace, "-9999999");
            


        }

        private void GetAndMoveFails()
        {
            //get all fails
            values = null;

            last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            range = worksheet.get_Range("A" + startRowReadExcel, last);

            values = (object[,])range.Value2;
            nRows = worksheet.UsedRange.Rows.Count;


            //sort to have fails on top of table
            range.Sort(range.Columns[7], Excel.XlSortOrder.xlAscending);


            //get index of the last fail
            var lastFailIndex = 0;
            for (int i = 1; i <= nRows - startRowReadExcel + 1; i++)
            {
                if (values[i, 7].ToString().Contains("F"))
                {
                    lastFailIndex++;
                }
            }
            lastFailIndex += startRowReadExcel - 1;


            //get fails range
            //Excel.Range lastIndexOfFailUnit = worksheet.Cells[lastFailIndex, nCols];
            //Excel.Range rangeFails = worksheet.get_Range("A" + startRowReadExcel, lastIndexOfFailUnit);

            last = worksheet.Cells[lastFailIndex, nCols];
            range = worksheet.get_Range("A" + startRowReadExcel, last);
            var rS = range;

            fails = (object[,])range.Value;


            //unsort 1 table
            //range.Sort(range.Columns[5], Excel.XlSortOrder.xlAscending);

            string currentSheet = "Unique Failed Data";
            var excelSheets = workbook.Worksheets;
            worksheet = excelSheets.get_Item(currentSheet);
            worksheet.Activate();


            //delete table in 2 worksheet
            last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            range = worksheet.get_Range("A" + startRowReadExcel, last);
            Excel.Range entireRow = range.EntireRow;
            entireRow.Delete(Excel.XlDirection.xlUp);

            app.ActiveWorkbook.SaveAs(openFileDialog1.FileName);


            //copy and paste fails in 2 worksheet
            last = worksheet.Cells[lastFailIndex, nCols];
            range = worksheet.get_Range("A" + startRowReadExcel, last);
            var rD = range;

            rS.Copy(rD);
            //range.AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormat3DEffects1, true, false, true, false, true, true);


        }

        private void CloseExcel()
        {
            app.ActiveWorkbook.Close();
            app.Quit();
            if (worksheet != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet);
            if (workbook != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
            if (app != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
        }

        private void CloseApp(object sender, FormClosingEventArgs e)
        {
            if (range != null)
            {
                app.ActiveWorkbook.Close();
                app.Quit();
                if (worksheet != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                if (app != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            }
        }

        private void ChangeValues()
        {
            for (int i = startRowReadExcel; i <= nRows; i++)
            {
                for (int j = startColumnReadExcel; j <= nCols; j++)
                {
                    if (values[i - (startRowReadExcel - 1), j] == null)
                    {
                        if (i != startRowReadExcel)
                            values[i - (startRowReadExcel - 1), j] = values[i - (startRowReadExcel - 1) - 1, j];
                    }
                }
            }
        }

        public void DeleteDuplicates(object[,] table)
        {
            int index = 0;

            for (int i = startRowReadExcel; i <= nRows - startRowReadExcel; i++)
            {
                if (table[i, 9].ToString() == table[i + 1, 9].ToString())
                {
                    last = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    range = worksheet.Cells[startRowReadExcel + i - 1 - index, nCols];
                    range.EntireRow.Delete(Type.Missing);
                    index++;
                }

            }
        }
    }
}

