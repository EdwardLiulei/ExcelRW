using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

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
            throw new NotImplementedException();
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            throw new NotImplementedException();
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
