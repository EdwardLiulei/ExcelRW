using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelRange:StdExcelRangeBase
    {
        #region field
        private ExcelRange _epRange;
        #endregion

        #region Constructor

        public EpExcelRange(ExcelRange range)
        {

            _epRange = range;
        }

        #endregion

        public override void SetFontStyle(System.Drawing.Font font)
        {
            throw new NotImplementedException();
        }

        public override void SetBold()
        {
            throw new NotImplementedException();
        }

        public override void SetItalic()
        {
            throw new NotImplementedException();
        }

        public override void UnBold()
        {
            throw new NotImplementedException();
        }

        public override void UnItalic()
        {
            throw new NotImplementedException();
        }

        public override void SetBackgroudColor(System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetFontColor(System.Drawing.Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetMerge()
        {
            throw new NotImplementedException();
        }

        public override void UnMerge()
        {
            throw new NotImplementedException();
        }
    }
}
