using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using  DataTable = System.Data.DataTable;
using System.Data;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiWorksheet:StdExcelWorkSheetBase
    {
        #region Field
        private ISheet _npoiWorsheet;

        #endregion

        #region Properity

        public override int ColumnNum
        {
            get
            {
                int columnNum = 0;
                int rowNum = _npoiWorsheet.LastRowNum;
                for (int i = 0; i < rowNum; i++)
                {
                    IRow row = _npoiWorsheet.GetRow(i);
                    if (row == null)
                        continue;
                    if (row.LastCellNum > columnNum)
                        columnNum = row.LastCellNum;
                }
                return columnNum;
            }
        }
        public override int RowNum
        {
            get
            {
                return _npoiWorsheet.LastRowNum;
            }
        }

        #endregion

        public NpoiWorksheet(ISheet worksheet)
        {
            _npoiWorsheet = worksheet;
            _sheetName = worksheet.SheetName;
        }

        public override string GetCellValue(int rowNumber, int columNumber)
        {
            if (rowNumber > _npoiWorsheet.LastRowNum)
                return "";
            if (columNumber > _npoiWorsheet.GetRow(rowNumber).LastCellNum)
                return "";
            ICell cell = _npoiWorsheet.GetRow(rowNumber).GetCell(columNumber);
            return GetCellValue(cell).ToString();
            //throw new NotImplementedException();
        }

        public override System.Data.DataTable GetTableContent(bool hasHeader = false)
        {
            bool isColumnName = true;
            int startRow = 1;
            ICell cell;
            DataColumn column;
            IRow row;
            System.Data.DataTable dataTable = new DataTable();
            if (_npoiWorsheet == null)
                return dataTable;

            int rowCount = _npoiWorsheet.LastRowNum;//总行数  
            if (rowCount > 0)
            {
                IRow firstRow = _npoiWorsheet.GetRow(0);//第一行  
                int cellCount = firstRow.LastCellNum;//列数  

                //构建datatable的列  
                if (isColumnName)
                {
                    startRow = 1;//如果第一行是列名，则从第二行开始读取  
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        cell = firstRow.GetCell(i);

                        string cellValue = GetCellValue(cell).ToString();
                        column = new DataColumn(cellValue);
                        dataTable.Columns.Add(column);

                    }
                }
                else
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        column = new DataColumn("column" + (i + 1));
                        dataTable.Columns.Add(column);
                    }
                }

                //填充行  
                for (int i = startRow; i <= rowCount; ++i)
                {
                    row = _npoiWorsheet.GetRow(i);
                    if (row == null) continue;

                    DataRow dataRow = dataTable.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        cell = row.GetCell(j);
                        if (cell == null)
                        {
                            dataRow[j] = "";
                        }
                        else
                        {
                            dataRow[j] = GetCellValue(cell);
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }



            return dataTable;
        }

        public object GetCellValue(ICell cell)
        {
            if (cell == null)
                return "";
            object value = null;
            try
            {
                if (cell.CellType != CellType.Blank)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            // Date comes here
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                value = cell.DateCellValue;
                            }
                            else
                            {
                                // Numeric type
                                value = cell.NumericCellValue;
                            }
                            break;
                        case CellType.Boolean:
                            // Boolean type
                            value = cell.BooleanCellValue;
                            break;
                        case CellType.Formula:
                            value = cell.CellFormula;
                            break;
                        default:
                            // String type
                            value = cell.StringCellValue;
                            break;
                    }
                }
            }
            catch (Exception)
            {
                value = "";
            }

            return value;
        }


        public override StdExcelRangeBase GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            return new NpoiExcelRange(_npoiWorsheet,startRow,startCol,endRow,endCol);
        }

        public override StdExcelCellBase GetCell(int rowNum, int columnNum)
        {
            rowNum -= 1;
            columnNum -= 1;
            IRow row = _npoiWorsheet.GetRow(rowNum);
            if (row == null)
                row = _npoiWorsheet.CreateRow(rowNum);

            ICell cell = row.GetCell(columnNum);
            if (cell == null)
                cell = row.CreateCell(columnNum);

            return new NpoiExcelCell(cell);
        }

        public override string GetCellFormula(int rowNum, int columnNum)
        {
            StdExcelCellBase cell = GetCell(rowNum, columnNum);
            return cell.GetValue();
        }

        public override StdExcelRowBase GetRow(int index)
        {
            return new NpoiExcelRow(_npoiWorsheet, index);
            
        }

        public override StdExcelColumnBase GetColumn(int index)
        {
            return new NpoiExcelColumn(_npoiWorsheet, index);
        }

        public override void InsertRow(int index)
        {
            _npoiWorsheet.ShiftRows(index + 1, _npoiWorsheet.LastRowNum,1);
            _npoiWorsheet.CreateRow(index);
        }

        public override void InsertColumn(int index)
        {

        }

        public override void SetCellValue(string value, int rowNum, int columnNum)
        {
         
            StdExcelCellBase cell = GetCell(rowNum, columnNum);
            cell.SetValue(value);
        }

        public override void SetCellFormula(string formular, int rowNum, int columnNum)
        {
            StdExcelCellBase cell = GetCell(rowNum, columnNum);
            cell.SetValue(formular);
        }

        public override void SetRangeColor(StdExcelRangeBase range, System.Drawing.Color color)
        {
            range.SetBackgroudColor(color);
        }

        public override void SetCellColor(int rowNum, int columnNum, System.Drawing.Color color)
        {
            StdExcelCellBase cell = GetCell(rowNum, columnNum);
            cell.SetBackgroudColor(color);

        }

        public override void MergeCell(StdExcelRangeBase range)
        {
            range.SetMerge();
        }

        public override void MergeCell(int startRow, int startCol, int endRow, int endCol)
        {
            StdExcelRangeBase range = GetRange(startRow-1, startCol-1, endRow-1, endCol-1);
            range.SetMerge();
        }

        public override List<string> GetSheetDataFromRow(int rowNum)
        {
            List<string> rowData = new List<string>();
            
            for (int i = 0; i < ColumnNum; i++)
            {
                string cellValue = GetCellValue(rowNum, i);
                rowData.Add(cellValue);

            }

            return rowData;
        }

        public override List<string> GetSheetDataFromColumn(int columnNum)
        {
            List<string> columnData = new List<string>();

            for (int i = 0; i < RowNum; i++)
            {
                string cellValue = GetCellValue(i, columnNum);
                columnData.Add(cellValue);

            }

            return columnData;
        }
    }
}
