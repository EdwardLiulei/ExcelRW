using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelRange:StdExcelRangeBase
    {
        #region field
        private ExcelRange _epRange;
        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                return _epRange.Style.Font.Bold;
            }
            set
            {
                _epRange.Style.Font.Bold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                return _epRange.Style.Font.Italic;
            }
            set
            {
                _epRange.Style.Font.Italic = value;
            }
        }

        #endregion

        #region Constructor

        public EpExcelRange(ExcelRange range)
        {

            _epRange = range;
        }

        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            _epRange.Style.Font.Name = font.Name;
            _epRange.Style.Font.Size = font.Size;
            _epRange.Style.Font.Italic = font.Italic;
            _epRange.Style.Font.Bold = font.Bold;
            _epRange.Style.Font.UnderLine = font.Underline;
        }

       

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            _epRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            _epRange.Style.Fill.BackgroundColor.SetColor(color);
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            _epRange.Style.Font.Color.SetColor(color);
        }

        public override void SetMerge()
        {
            _epRange.Merge = true;
        }

        public override void UnMerge()
        {
            _epRange.Merge = false;
        }

        public override string[,] GetRangeData()
        {
            throw new NotImplementedException();
        }
    }
}
