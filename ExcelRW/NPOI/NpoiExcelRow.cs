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
    public class NpoiExcelRow:StdExcelRowBase
    {
        #region Field
        private IRow _row;
        private int _rowNum;
        private ISheet _npoiWorksheet;

        #endregion

        #region Porperity

        public override bool Bold
        {
            get
            {
                var workbook = _npoiWorksheet.Workbook;
                return _row.RowStyle.GetFont(workbook).IsBold;
            }
            set
            {
                var workbook = _npoiWorksheet.Workbook;
                _row.RowStyle.GetFont(workbook).IsBold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                var workbook = _npoiWorksheet.Workbook;
                return _row.RowStyle.GetFont(workbook).IsItalic;
            }
            set
            {
                var workbook = _npoiWorksheet.Workbook;
                _row.RowStyle.GetFont(workbook).IsItalic = value;
            }
        }
        #endregion

        #region Constructor
        public NpoiExcelRow(ISheet sheet,int rowNum)
        {
            _row = sheet.GetRow(rowNum);
            _rowNum = rowNum;
            _npoiWorksheet = sheet;
        }

        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            var workbook = _npoiWorksheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(_row.RowStyle);

            IFont thisFont = cellStyle.GetFont(workbook);
            thisFont.FontName = font.Name;
            thisFont.IsBold = font.Bold;
            thisFont.IsItalic = font.Italic;
            if (font.Underline)
                thisFont.Underline = FontUnderlineType.Single;

            cellStyle.SetFont(thisFont);

            _row.RowStyle = cellStyle;
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            IWorkbook workbook = _npoiWorksheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(_row.RowStyle);
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
            _row.RowStyle = cellStyle;
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            IWorkbook workbook = _npoiWorksheet.Workbook;
            ICellStyle cellStyle = workbook.CreateCellStyle();
            if (workbook is HSSFWorkbook)
            {
                HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                HSSFPalette palette = hssfWorkbook.GetCustomPalette(); //调色板实例

                //palette.SetColorAtIndex((short)8, color.R, color.G, color.B);

                HSSFColor hssFColor = palette.FindSimilarColor(color.R, color.G, color.B);
                cellStyle.CloneStyleFrom(_row.RowStyle);

                IFont font = cellStyle.GetFont(workbook);
                font.Color = hssFColor.Indexed;
                cellStyle.SetFont(font);
                _row.RowStyle = cellStyle;

            }
            else
            {

                //No way!
            }
        }

        public override void SetHeight(int height)
        {
            _row.Height = (short)height;
        }

        public override StdExcelCellBase GetCell(int columnNum)
        {
            ICell cell = _row.GetCell(columnNum);
            return new NpoiExcelCell(cell);
        }
    }
}
