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

        #region Properity
        public override bool Bold
        {
            get
            {
                var workbook = _npoiCell.Sheet.Workbook;
                return _npoiCell.CellStyle.GetFont(workbook).IsBold;
            }
            set
            {
                var workbook = _npoiCell.Sheet.Workbook;
                _npoiCell.CellStyle.GetFont(workbook).IsBold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                var workbook = _npoiCell.Sheet.Workbook;
                return _npoiCell.CellStyle.GetFont(workbook).IsItalic;
            }
            set
            {
                var workbook = _npoiCell.Sheet.Workbook;
                _npoiCell.CellStyle.GetFont(workbook).IsItalic = value;
            }
        }
        #endregion

        #region Constructor
        public NpoiExcelCell(ICell cell)
        {
            _npoiCell = cell; 
        }

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
            var workbook = _npoiCell.Sheet.Workbook;
            IFont thisFont = _npoiCell.CellStyle.GetFont(workbook);
            thisFont.FontName = font.Name;
            thisFont.IsBold = font.Bold;
            thisFont.IsItalic = font.Italic;
            if (font.Underline)
                thisFont.Underline = FontUnderlineType.Single;

            _npoiCell.CellStyle.SetFont(thisFont);

        }
   

        public override void SetBackgroudColor(Color color)
        {
            IWorkbook workbook = _npoiCell.Sheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(_npoiCell.CellStyle);
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                HSSFPalette palette = hssfWorkbook.GetCustomPalette(); //调色板实例

                //palette.SetColorAtIndex((short)8, color.R, color.G, color.B);

                HSSFColor hssFColor = palette.FindSimilarColor(color.R, color.G, color.B);

                cellStyle.FillPattern = FillPattern.SolidForeground;

                cellStyle.FillForegroundColor = hssFColor.Indexed;
                
            }
            else
            {
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
                HSSFPalette palette = hssfWorkbook.GetCustomPalette(); //调色板实例

                //palette.SetColorAtIndex((short)8, color.R, color.G, color.B);

                HSSFColor hssFColor = palette.FindSimilarColor(color.R, color.G, color.B);

                cellStyle.FillPattern = FillPattern.SolidForeground;

                cellStyle.FillForegroundColor = hssFColor.Indexed;
                //No way!
            }
            _npoiCell.CellStyle = cellStyle;

        }

       
        public override void SetFontColor(Color color)
        {
            IWorkbook workbook = _npoiCell.Sheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                HSSFPalette palette = hssfWorkbook.GetCustomPalette(); //调色板实例

                //palette.SetColorAtIndex((short)8, color.R, color.G, color.B);

                HSSFColor hssFColor = palette.FindSimilarColor(color.R, color.G, color.B);
                cellStyle.CloneStyleFrom(_npoiCell.CellStyle);

                IFont font = cellStyle.GetFont(workbook);
                font.Color = hssFColor.Indexed;
                cellStyle.SetFont(font);
                _npoiCell.CellStyle = cellStyle;

            }
            else
            {
               
                //No way!
            }
        }

       
    }
}
