using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using OfficeOpenXml;

namespace ExcelReadAndWrite.Epplus
{
    public class EpExcelRow:StdExcelRowBase
    {
        #region Field
        private ExcelRow _epRow;
        #endregion

        #region Constructor
        public EpExcelRow(ExcelRow row)
        {
            _epRow = row;
 
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

        public override void SetHeight(int height)
        {
            throw new NotImplementedException();
        }

        public override StdExcelCellBase GetCell(int columnNum)
        {
            throw new NotImplementedException();
        }
    }
}
