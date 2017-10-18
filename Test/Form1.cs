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
            Thread newt = new Thread(NewThread);
            newt.Start();
           
            
        }

        private void NewThread()
        {
            Thread.Sleep(5000);
            MessageBox.Show("ok");
        }

    }
}
