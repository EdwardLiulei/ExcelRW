using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelRow:StdExcelRowBase
    {
        #region Field
        private IRow _row;
        private int _rowNum;

        #endregion

        #region Porperity

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
        #endregion

        #region Constructor
        public NpoiExcelRow(ISheet sheet,int rowNum)
        {
            _row = sheet.GetRow(rowNum);
            _rowNum = rowNum;
        }

        #endregion

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

        public override void SetHeight(int height)
        {
            throw new NotImplementedException();
        }

        public override StdExcelCellBase GetCell(int columnNum)
        {
            ICell cell = _row.GetCell(columnNum);
            return new NpoiExcelCell(cell);
        }
    }
}
