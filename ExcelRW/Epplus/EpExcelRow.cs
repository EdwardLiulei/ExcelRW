﻿using System;
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
        private int _rowNum;
        private ExcelWorksheet _workSheet;
        #endregion

        #region Properity

        public override bool Bold
        {
            get
            {
                return _epRow.Style.Font.Bold;
            }
            set
            {
                _epRow.Style.Font.Bold = value;
            }
        }

        public override bool Italic
        {
            get
            {
                return _epRow.Style.Font.Italic;
            }
            set
            {
                _epRow.Style.Font.Italic = value;
            }
        }

        #endregion

        #region Constructor
        public EpExcelRow(ExcelWorksheet sheet, int rowNum)
        {
            _workSheet = sheet;
            _rowNum = rowNum;
            _epRow = _workSheet.Row(rowNum);
 
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
            throw new NotImplementedException();
        }
    }
}
