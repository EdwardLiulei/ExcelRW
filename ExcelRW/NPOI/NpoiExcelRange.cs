using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelRange:StdExcelRangeBase
    {
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

        public override void SetFontStyle(System.Drawing.Font font)
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

        public override string[,] GetRangeData()
        {
            throw new NotImplementedException();
        }
    }
}
