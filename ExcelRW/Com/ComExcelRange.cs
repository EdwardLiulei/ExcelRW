using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ExcelReadAndWrite.Com
{
    public class ComExcelRange : StdExcelRangeBase
    {

        #region Field
        private Range _comRange;
        #endregion
        public override void SetBackgroudColor(Color color)
        {
            
        }

        public override void SetBold()
        {
            throw new NotImplementedException();
        }

        public override void SetFontColor(Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetFontStyle(System.Drawing.Font font)
        {
            throw new NotImplementedException();
        }


        public override void SetItalic()
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
