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
        private ExcelWorksheet _epWorkSheet;
        private int _columnNum;

        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                return _epColumn.Style.Font.Bold;
            }
            set
            {
                _epColumn.Style.Font.Bold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                return _epColumn.Style.Font.Italic;
            }
            set
            {
                _epColumn.Style.Font.Italic = value;
            }
        }

        #endregion

        #region Constructor

        public EpExcelColumn(ExcelWorksheet sheet,int columnNum)
        {
            _epWorkSheet = sheet;
            _columnNum = columnNum;
            _epColumn = sheet.Column(columnNum);
 
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
            ExcelRange range = _epWorkSheet.Cells[rowNum, _columnNum];
            return new EpExcelCell(range);
        }
    }
}
