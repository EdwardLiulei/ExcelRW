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
                IRow row = _npoiWorksheet.GetRow(i);
                ICell cell = row.GetCell(_columnNum);

                IFont thisFont = cell.CellStyle.GetFont(workbook);
                thisFont.FontName = font.Name;
                thisFont.IsBold = font.Bold;
                thisFont.IsItalic = font.Italic;
                if (font.Underline)
                    thisFont.Underline = FontUnderlineType.Single;

                cell.CellStyle.SetFont(thisFont);
            }
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetWidth(int width)
        {
            throw new NotImplementedException();
        }

        public override StdExcelCellBase GetCell(int rowNum)
        {
            throw new NotImplementedException();
        }
    }
}
