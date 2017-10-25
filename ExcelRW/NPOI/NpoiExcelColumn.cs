using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelColumn:StdExcelColumnBase
    {

        #region Field

        private ISheet _npoiWorksheet;
        private int _columnNum;

        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                var workbook = _npoiWorksheet.Workbook;
                return _npoiWorksheet.GetColumnStyle(_columnNum).GetFont(workbook).IsBold;
            }
            set
            {
                var workbook = _npoiWorksheet.Workbook;
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.CloneStyleFrom(_npoiWorksheet.GetColumnStyle(_columnNum));
                cellStyle.GetFont(workbook).IsBold = value;
                _npoiWorksheet.SetDefaultColumnStyle(_columnNum, cellStyle);
            }
        }

        public override bool Italic
        {
            get
            {
                var workbook = _npoiWorksheet.Workbook;
               
                return _npoiWorksheet.GetColumnStyle(_columnNum).GetFont(workbook).IsItalic;

            }
            set
            {
                var workbook = _npoiWorksheet.Workbook;
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.CloneStyleFrom(_npoiWorksheet.GetColumnStyle(_columnNum));
                cellStyle.GetFont(workbook).IsItalic = value;
                _npoiWorksheet.SetDefaultColumnStyle(_columnNum, cellStyle);
            }
        }

        #endregion

        #region Constructor
        public NpoiExcelColumn(ISheet sheet, int columnNum)
        {
            _npoiWorksheet = sheet;
            _columnNum = columnNum;
 
        }

        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            var workbook = _npoiWorksheet.Workbook;
            int rowNum = _npoiWorksheet.LastRowNum;
            for (int i = 0; i <= rowNum; i++)
            {
                ICellStyle cellStyle = workbook.CreateCellStyle();
                IRow row = _npoiWorksheet.GetRow(i);
                ICell cell = row.GetCell(_columnNum);

                cellStyle.CloneStyleFrom(cell.CellStyle);
                IFont thisFont = cellStyle.GetFont(workbook);
                thisFont.FontName = font.Name;
                thisFont.IsBold = font.Bold;
                thisFont.IsItalic = font.Italic;
                if (font.Underline)
                    thisFont.Underline = FontUnderlineType.Single;

                cellStyle.SetFont(thisFont);
                cell.CellStyle = cellStyle;
            }
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            IWorkbook workbook = _npoiWorksheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(_npoiWorksheet.GetColumnStyle(_columnNum));
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
            _npoiWorksheet.SetDefaultColumnStyle(_columnNum,cellStyle);
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            IWorkbook workbook = _npoiWorksheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(_npoiWorksheet.GetColumnStyle(_columnNum));
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                HSSFPalette palette = hssfWorkbook.GetCustomPalette(); //调色板实例

                //palette.SetColorAtIndex((short)8, color.R, color.G, color.B);

                HSSFColor hssFColor = palette.FindSimilarColor(color.R, color.G, color.B);
                IFont font = cellStyle.GetFont(workbook);
                font.Color = hssFColor.Indexed;
                cellStyle.SetFont(font);

            }
            else
            {
                
                //No way!
            }
            _npoiWorksheet.SetDefaultColumnStyle(_columnNum, cellStyle);

        }

        public override void SetWidth(int width)
        {
            _npoiWorksheet.SetColumnWidth(_columnNum,width);
        }

        public override StdExcelCellBase GetCell(int rowNum)
        {
            ICell cell = _npoiWorksheet.GetRow(rowNum).GetCell(_columnNum);
            return new NpoiExcelCell(cell);
        }
    }
}
