using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelRow:StdExcelRowBase
    {
        #region Field
        private ExcelRow _epRow;
        private int _rowNum;
        private ExcelWorksheet _workSheet;
        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                return _epRow.Style.Font.Bold;
            }
            set
            {
                _epRow.Style.Font.Bold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                return _epRow.Style.Font.Italic;
            }
            set
            {
                _epRow.Style.Font.Italic = value;
            }
        }

        #endregion

        #region Constructor
        public EpExcelRow(ExcelWorksheet sheet, int rowNum)
        {
            _workSheet = sheet;
            _rowNum = rowNum;
            _epRow = _workSheet.Row(rowNum);
 
        }

        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            _epRow.Style.Font.Name = font.Name;
            _epRow.Style.Font.Size = font.Size;
            _epRow.Style.Font.Italic = font.Italic;
            _epRow.Style.Font.Bold = font.Bold;
            _epRow.Style.Font.UnderLine = font.Underline;
        }

       

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _epRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            _epRow.Style.Fill.BackgroundColor.SetColor(color);
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _epRow.Style.Font.Color.SetColor(color);
        }

        public override void SetHeight(int height)
        {
            _epRow.Height = height;
        }

        public override StdExcelCellBase GetCell(int columnNum)
        {
            ExcelRange cell = _workSheet.Cells[_rowNum, columnNum];
            return new EpExcelCell(cell);
        }
    }
}
