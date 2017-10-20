using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using NPOI.HSSF.Util;
using NPOI.SS.Util;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelCell:StdExcelCellBase
    {
        #region Field
        private ICell _npoiCell;

        #endregion

        public override string GetValue()
        {
            ICell cell = _npoiCell;
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

        public override void SetValue(string value)
        {
            _npoiCell.SetCellValue(value);
        }

        public override void SetFormular(string formular)
        {
            _npoiCell.SetCellFormula(formular);
        }

        public override void SetFontStyle(System.Drawing.Font font)
        {
            //_npoiCell.CellStyle.SetFont(font);

        }

        public override void SetBold()
        {
            IWorkbook parentWorkbook = _npoiCell.Sheet.Workbook;
            _npoiCell.CellStyle.GetFont(parentWorkbook).IsBold = !_npoiCell.CellStyle.GetFont(parentWorkbook).IsBold;
        }

        public override void SetItalic()
        {
            IWorkbook parentWorkbook = _npoiCell.Sheet.Workbook;
            _npoiCell.CellStyle.GetFont(parentWorkbook).IsItalic = !_npoiCell.CellStyle.GetFont(parentWorkbook).IsItalic;
        }

        public override void SetBackgroudColor(Color color)
        {
            
            
        }

        public override void SetFontColor(Color color)
        {
            
        }
    }
}
