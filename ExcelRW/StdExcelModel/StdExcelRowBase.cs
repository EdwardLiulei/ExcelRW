using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel.BaseModel;
using System.Drawing;

namespace ExcelReadAndWrite.StdExcelModel
{
    public abstract class StdExcelRowBase:IStdRow
    {
        #region Field

        private int _rowNumber;

        #endregion

        #region Properity

        public abstract bool Bold { set; get; }

        public abstract bool Italic { set; get; }
        #endregion


        #region Abstarct Functions

        public abstract void SetFontStyle(Font font);

        public abstract void SetBackgroudColor(Color color);

        public abstract void SetFontColor(Color color);

        public abstract void SetHeight(int height);

        public abstract StdExcelCellBase GetCell(int columnNum);

        #endregion
    }
}
