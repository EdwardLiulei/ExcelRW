using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelReadAndWrite.StdExcelModel;
using System.Data;
using ExcelReadAndWrite.Util;


namespace ExcelReadAndWrite.Com
{
    public class ComWorksheet:StdExcelWorkSheetBase
    {
        #region Field
        private Worksheet _worksheet;

        #endregion

        #region Constructor
        public ComWorksheet(Worksheet worksheet)
        {
            _worksheet = worksheet;

            _sheetName = worksheet.Name;
        }

        #endregion

        public Worksheet GetComWorksheet()
        {
            return _worksheet;
        }

        public override System.Data.DataTable GetTableContent(bool hasHeader = false)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _worksheet.UsedRange.Rows.Count;
            int iColCount = _worksheet.UsedRange.Columns.Count;
            object[,] a = new string[iRowCount,iColCount];
            
            a =  _worksheet.Range[_worksheet.Cells[1,1],_worksheet.Cells[iRowCount,iColCount]].Value2;
           
            if (hasHeader == true)
            {
                List<string> headers = new List<string>();
                for (int i = 0; i < iColCount; i++)
                {
                    headers.Add(a[0,i].ToString());
                }
                if (headers.Count > headers.Distinct().Count())
                    throw new Exception(string.Format("The sheet: {0} contains duplicate headers",_sheetName));
                foreach (var header in headers)
                {
                    table.Columns.Add(header);
                }
            }
            else
            {
                for (int i = 0; i < iColCount; i++)
                {
                    string columnName = WorksheetAddress.GetColumnAddress(i + 1);
                    table.Columns.Add(columnName);
                }
                DataRow row = table.NewRow();
                for (int j = 1; j < iColCount; j++)
                {
                    row[j] = a[1, j];
                }
                table.Rows.Add(row);
            }


            for (int i = 2; i < iRowCount; i++)
            {
                DataRow row = table.NewRow();
                for (int j = 1; j < iColCount; j++)
                {
                    if (a[i, j] == null)
                        row[j] = "";
                    else
                        row[j] = a[i, j].ToString();
                }
                table.Rows.Add(row);
            }
            return table;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            Range rang= _worksheet.Cells[rowNumber, columNumber];
            return rang.Value;
        }


        public override StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            return new ComExcelRange(_worksheet.Range[_worksheet.Cells[startRow,startCol],_worksheet.Cells[endRow,endCol]]);
        }

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            Range cell = _worksheet.Cells[rowNum, columnNum];
            return new ComExcelCell(cell);
        }

        public override string GetCellFormula(int rowNumber, int columNumber)
        {
            Range rang = _worksheet.Cells[rowNumber, columNumber];
            return rang.Formula;
            
        }

        public override StdExcelRowBase GetRow(int index)
        {
            Range row = _worksheet.Rows[index];
            return new ComExcelRow(row);
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            Range column = _worksheet.Columns[index];
            return new ComExcelColumn(column);
        }

        public override void InsertRow(int index)
        {
            _worksheet.Rows.Insert(index);
        }

        public override void InsertColumn(int index)
        {
            _worksheet.Columns.Insert(index);
        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        {
            GetCell(rowNum, columnNum).SetValue(value);
        }

        public override void SetCellFormula(string formula, int rowNum, int columnNum)
        {
            GetCell(rowNum, columnNum).SetFormular(formula);
        }

        public override void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color)
        {
            range.SetBackgroudColor(color);
        }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        {
            GetCell(rowNum, columnNum).SetBackgroudColor(color);
        }

        public override void MergeCell(StdExcelRangeBase range) 
        {
            range.SetMerge();
        }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol) 
        {
            GetRange(startRow, startCol, endRow, endCol).SetMerge();
        }
    }
}
