using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;


namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelRange:StdExcelRangeBase
    {
        #region Field
        private ISheet _npoiWorksheet;
        private CellRangeAddress _rangeAddress;

        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public override bool Italic
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        #endregion

        #region Constructor

        public NpoiExcelRange(ISheet sheet,int startRow,int startColumn,int endRow,int endColumn)
        {
            _npoiWorksheet = sheet;
            _rangeAddress = new CellRangeAddress(startRow, endRow, startColumn, endColumn);
 
        }

        #endregion
        public override void SetFontStyle(System.Drawing.Font font)
        {
            for (int i = _rangeAddress.FirstRow; i <= _rangeAddress.LastRow; i++)
            {
                IRow row = _npoiWorksheet.GetRow(i);
                if (row == null)
                    row = _npoiWorksheet.CreateRow(i);
                for (int j = _rangeAddress.FirstColumn; j <= _rangeAddress.LastColumn; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                        cell = row.CreateCell(j);
                    NpoiExcelCell excelCell = new NpoiExcelCell(cell);
                    excelCell.SetFontStyle(font);
                }
            }
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            for (int i = _rangeAddress.FirstRow; i <= _rangeAddress.LastRow; i++)
            {
                IRow row = _npoiWorksheet.GetRow(i);
                if (row == null)
                    row = _npoiWorksheet.CreateRow(i);
                for (int j = _rangeAddress.FirstColumn; j <= _rangeAddress.LastColumn; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                        cell = row.CreateCell(j);
                    NpoiExcelCell excelCell = new NpoiExcelCell(cell);
                    excelCell.SetBackgroudColor(color);
                }
            }
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            for (int i = _rangeAddress.FirstRow; i <= _rangeAddress.LastRow; i++)
            {
                IRow row = _npoiWorksheet.GetRow(i);
                if (row == null)
                    row = _npoiWorksheet.CreateRow(i);
                for (int j = _rangeAddress.FirstColumn; j <= _rangeAddress.LastColumn; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (cell == null)
                        cell = row.CreateCell(j);
                    NpoiExcelCell excelCell = new NpoiExcelCell(cell);
                    excelCell.SetFontColor(color);
                }
            }
        }

        public override void SetMerge()
        {
            _npoiWorksheet.AddMergedRegion(_rangeAddress);
        }

        public override void UnMerge()
        {
            int mergeCount = _npoiWorksheet.NumMergedRegions;
            for (int i = mergeCount - 1; i >= 0; i--)
            {
                var range = _npoiWorksheet.GetMergedRegion(i);
                if (range.FirstRow == _rangeAddress.FirstRow && range.FirstColumn == _rangeAddress.FirstColumn &&
                    range.LastColumn == _rangeAddress.LastColumn && range.LastRow == _rangeAddress.LastRow)
                    _npoiWorksheet.RemoveMergedRegion(i);
            }
        }

        public override string[,] GetRangeData()
        {
            int columnNum = _rangeAddress.LastColumn - _rangeAddress.FirstColumn + 1;
            int rowNum = _rangeAddress.LastRow - _rangeAddress.FirstRow + 1;
            string[,] values = new string[rowNum, columnNum];
            for (int i = 0; i < rowNum; i++)
                for (int j = 0; j < columnNum; j++)
                {
                    IRow row = _npoiWorksheet.GetRow(i+ _rangeAddress.FirstRow);
                    if (row == null)
                    {
                        values[i, j] = "";
                    }
                    else
                    {
                        values[i, j] = GetValue(row.GetCell(j+ _rangeAddress.FirstColumn));
                    }
                }

            return values;
        }

        private string GetValue(ICell cell)
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

            return value.ToString();

        }
    }
}
