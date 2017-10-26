using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlExcelRange : StdExcelRangeBase
    {
        public override bool Bold { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override bool Italic { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public override string[,] GetRangeData()
        {
            throw new NotImplementedException();
        }

        public override void SetBackgroudColor(Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetFontColor(Color color)
        {
            throw new NotImplementedException();
        }

        public override void SetFontStyle(Font font)
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
