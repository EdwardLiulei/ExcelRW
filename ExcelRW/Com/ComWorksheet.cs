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
        private Worksheet _comWorksheet;

        #endregion

        #region Properity
        public override int ColumnNum
        {
            get { return _comWorksheet.UsedRange.Columns.Count; }
        }

        public override int RowNum
        {
            get { return _comWorksheet.UsedRange.Rows.Count; }
        }

        #endregion

        #region Constructor
        public ComWorksheet(Worksheet worksheet)
        {
            _comWorksheet = worksheet;

            _sheetName = worksheet.Name;
        }

        #endregion

        #region Public Functions
        public Worksheet GetComWorksheet()
        {
            return _comWorksheet;
        }

        public override System.Data.DataTable GetTableContent(bool hasHeader = false)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            int iRowCount = _comWorksheet.UsedRange.Rows.Count;
            int iColCount = _comWorksheet.UsedRange.Columns.Count;
            object[,] a = new string[iRowCount,iColCount];
            
            a =  _comWorksheet.Range[_comWorksheet.Cells[1,1],_comWorksheet.Cells[iRowCount,iColCount]].Value2;
           
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
            Range rang= _comWorksheet.Cells[rowNumber, columNumber];
            return rang.Value;
        }

        public override StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            return new ComExcelRange(_comWorksheet.Range[_comWorksheet.Cells[startRow,startCol],_comWorksheet.Cells[endRow,endCol]]);
        }

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            Range cell = _comWorksheet.Cells[rowNum, columnNum];
            return new ComExcelCell(cell);
        }

        public override string GetCellFormula(int rowNumber, int columNumber)
        {
            Range rang = _comWorksheet.Cells[rowNumber, columNumber];
            return rang.Formula;
            
        }

        public override StdExcelRowBase GetRow(int index)
        {
            Range row = _comWorksheet.Rows[index];
            return new ComExcelRow(row);
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            Range column = _comWorksheet.Columns[index];
            return new ComExcelColumn(column);
        }

        public override void InsertRow(int index)
        {
            _comWorksheet.Rows.Insert(index);
        }

        public override void InsertColumn(int index)
        {
            _comWorksheet.Columns.Insert(index);
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

        public override List<string> GetSheetDataFromRow(int rowNum)
        {
            List<string> rowData = new List<string>();
            int columnNum = _comWorksheet.Columns.Count;
            string[,] dataArray = new string[1, columnNum];
            dataArray = _comWorksheet.Range[_comWorksheet.Cells[rowNum,1],_comWorksheet.Cells[rowNum,columnNum]].Value2;
            for (int i = 0; i < columnNum; i++)
            {
                rowData.Add(dataArray[1, i]);
            }

            return rowData;
        }

        public override List<string> GetSheetDataFromColumn(int columnNum)
        {
            List<string> columnData = new List<string>();
            int rowNum = _comWorksheet.Rows.Count;
            string[,] dataArray = new string[rowNum, 1];
            dataArray = _comWorksheet.Range[_comWorksheet.Cells[rowNum, 1], _comWorksheet.Cells[rowNum, columnNum]].Value2;
            for (int i = 0; i < columnNum; i++)
            {
                columnData.Add(dataArray[i, 1]);
            }

            return columnData;
        }

        #endregion
    }
}
