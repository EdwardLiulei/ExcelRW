using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComExcelColumn:StdExcelColumnBase
    {
        #region Field
        private Range _comColumn;
        #endregion

        #region Constructor
        public ComExcelColumn(Range range)
        {
            _comColumn = range;
 
        }

        #endregion
        public override void SetFontStyle(System.Drawing.Font font)
        {
            _comColumn.Font.Name = font.Name;
            _comColumn.Font.Size = font.Size;
            _comColumn.Font.Bold = font.Bold;
            _comColumn.Font.Italic = font.Italic;
            _comColumn.Font.Underline = font.Underline;
        }

        public override void SetBold()
        {
            _comColumn.Font.Bold = true;
        }

        public override void SetItalic()
        {
            _comColumn.Font.Italic = true;
        }

        public override void UnBold()
        {
            _comColumn.Font.Bold = false;
        }

        public override void UnItalic()
        {
            _comColumn.Font.Italic = false;
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _comColumn.Interior.Color = System.Drawing.Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _comColumn.Font.Color = System.Drawing.Color.FromArgb(color.A, color.B, color.G, color.R).ToArgb();
        }

        public override void SetWidth(int width)
        {
            _comColumn.ColumnWidth = width;
        }

        public override StdExcelCellBase GetCell(int rowNum)
        {
            Range range = _comColumn.Cells[rowNum];
            return new ComExcelCell(range);
        }
    }
}
