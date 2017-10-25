using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;


namespace ExcelReadAndWrite.NPOI
{
    public class NpoiExcelRange:StdExcelRangeBase
    {
        #region Field
        private ISheet _npoiWorksheet;
        private CellRangeAddress _rangeAddress;

        #endregion

        #region Properity

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

        public NpoiExcelRange(ISheet sheet,int startRow,int startColumn,int endRow,int endColumn)
        {
            _npoiWorksheet = sheet;
            _rangeAddress = new CellRangeAddress(startRow, endRow, startColumn, endColumn);
 
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

        public override void SetMerge()
        {
            _npoiWorksheet.AddMergedRegion(_rangeAddress);
        }

        public override void UnMerge()
        {
            int mergeCount = _npoiWorksheet.NumMergedRegions;
            for (int i = mergeCount - 1; i >= 0; i--)
            {
                var range = _npoiWorksheet.GetMergedRegion(i);
                if (range.FirstRow == _rangeAddress.FirstRow && range.FirstColumn == _rangeAddress.FirstColumn &&
                    range.LastColumn == _rangeAddress.LastColumn && range.LastRow == _rangeAddress.LastRow)
                    _npoiWorksheet.RemoveMergedRegion(i);
            }
        }

        public override string[,] GetRangeData()
        {
            throw new NotImplementedException();
        }
    }
}
