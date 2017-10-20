using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using ExcelReadAndWrite.Com;
using ExcelReadAndWrite.Epplus;
using ExcelReadAndWrite.StdExcelModel;
using ExcelReadAndWrite.Oledb;
using ExcelReadAndWrite.NPOI;
using ExcelReadAndWrite.StdExcelModel.BaseModel;

namespace ExcelReadAndWrite
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object fileName = @"D:\Projects\ExcelRW\test.xls";
            if (!File.Exists(fileName.ToString()))
            {
                MessageBox.Show("Path not exists");
                return;
            }
            ParameterizedThreadStart ParStart = new ParameterizedThreadStart(Read);
            Thread newt = new Thread(ParStart);
            newt.Start(fileName);
           
            
        }



        private void Read(object fileName)
        {
            IStdWorkbook workbook = new ComWorkbook();
            workbook.Load(fileName.ToString());
            
            var t = workbook.WorkSheets.First().GetTableContent();
            MessageBox.Show("ok");
            //workbook.ReleaseReSource();
           
        }

    }
}
