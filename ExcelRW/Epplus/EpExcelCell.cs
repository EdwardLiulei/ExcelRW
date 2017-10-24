using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelCell:StdExcelCellBase
    {
        #region Field
        private ExcelRange _epCell;
        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                return _epCell.Style.Font.Bold;
            }
            set
            {
                _epCell.Style.Font.Bold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                return _epCell.Style.Font.Italic;
            }
            set
            {
                _epCell.Style.Font.Italic = value;
            }
        }

        #endregion

        #region Constructor

        public EpExcelCell(ExcelRange range)
        {
            _epCell = range;
        }

        #endregion

        public override string GetValue()
        {
            if (_epCell.Value == null)
                return "";
            return _epCell.Value.ToString();
        }

        public override void SetValue(string value)
        {
            _epCell.Value = value;
        }

        public override void SetFormular(string formular)
        {
            _epCell.Formula = formular;
        }

        public override void SetFontStyle(System.Drawing.Font font)
        {
            _epCell.Style.Font.Name = font.Name;
            _epCell.Style.Font.Size = font.Size;
            _epCell.Style.Font.Italic = font.Italic;
            _epCell.Style.Font.Bold = font.Bold;
            _epCell.Style.Font.UnderLine = font.Underline;
        }
 

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _epCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            _epCell.Style.Fill.BackgroundColor.SetColor(color);
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _epCell.Style.Font.Color.SetColor(color);
        }
    }
}
