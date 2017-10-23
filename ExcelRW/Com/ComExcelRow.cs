using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComExcelRow:StdExcelRowBase
    {
        #region Field
        private Range _comRow;
        #endregion

        #region Constructor
        public ComExcelRow(Range range)
        {
            _comRow = range;
 
        }

        #endregion
        public override void SetFontStyle(System.Drawing.Font font)
        {

            _comRow.Font.Name = font.Name;
            _comRow.Font.Size = font.Size;
            _comRow.Font.Bold = font.Bold;
            _comRow.Font.Italic = font.Italic;
            _comRow.Font.Underline = font.Underline;
        }

        public override void SetBold()
        {
            _comRow.Font.Bold = true;
        }

        public override void SetItalic()
        {
            _comRow.Font.Italic= true;
        }

        public override void UnBold()
        {
            _comRow.Font.Bold = false;
        }

        public override void UnItalic()
        {
            _comRow.Font.Italic = false;
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _comRow.Interior.Color = color;
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _comRow.Font.Color = color;
        }

        public override void SetHeight(int height)
        {
            _comRow.RowHeight = height;
        }

        public override StdExcelCellBase GetCell(int columnNum)
        {
            Range cell = _comRow.Cells[columnNum];
            return new ComExcelCell(cell);
        }
    }
}
