using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlExcelColumn : StdExcelColumnBase
    {
       
        public override StdExcelCellBase GetCell(int rowNum)
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

        public override void SetWidth(int width)
        {
            throw new NotImplementedException();
        }

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
    }
}
