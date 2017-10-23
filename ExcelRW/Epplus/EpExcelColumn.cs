using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelColumn:StdExcelColumnBase
    {
        #region Field
        private ExcelColumn _epColumn;

        #endregion

        #region Constructor

        public EpExcelColumn(ExcelColumn column)
        {
            _epColumn = column;
 
        }
        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            _epColumn.Style.Font.Name = font.Name;
            _epColumn.Style.Font.Size = font.Size;
            _epColumn.Style.Font.Italic = font.Italic;
            _epColumn.Style.Font.Bold = font.Bold;
            _epColumn.Style.Font.UnderLine = font.Underline;
        }

        public override void SetBold()
        {
            _epColumn.Style.Font.Bold = true;
        }

        public override void SetItalic()
        {
            _epColumn.Style.Font.Italic = true;
        }

        public override void UnBold()
        {
            _epColumn.Style.Font.Bold = false;
        }

        public override void UnItalic()
        {
            _epColumn.Style.Font.Italic = false;
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _epColumn.Style.Fill.PatternType = ExcelFillStyle.Solid;
            _epColumn.Style.Fill.BackgroundColor.SetColor(color);
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _epColumn.Style.Font.Color.SetColor(color);
        }

        public override void SetWidth(int width)
        {
            _epColumn.Width = width;
        }

        public override StdExcelCellBase GetCell(int rowNum)
        {
            throw new NotImplementedException();
        }
    }
}
