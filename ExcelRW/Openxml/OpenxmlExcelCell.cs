using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelReadAndWrite.StdExcelModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelReadAndWrite.Openxml
{
    public class OpenxmlExcelCell : StdExcelCellBase
    {

        #region Field
        private Cell _opemxmlCell;

        #endregion

        #region Properity


        #endregion

        #region Constructor
        public OpenxmlExcelCell()
        {
        }
        #endregion

        public override string GetValue()
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

        public override void SetFontStyle(System.Drawing.Font font)
        {
            throw new NotImplementedException();
        }

        public override void SetFormular(string formular)
        {
            throw new NotImplementedException();
        }

        public override void SetValue(string value)
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
