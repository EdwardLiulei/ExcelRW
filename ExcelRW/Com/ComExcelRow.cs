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
        #region Properity

        public override bool Bold
        {
            set
            { _comRow.Font.Bold = value; }
            get
            { return _comRow.Font.Bold; }
        }

        public override bool Italic
        {
            set
            { _comRow.Font.Italic = value; }
            get
            { return _comRow.Font.Italic; }
        }

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
