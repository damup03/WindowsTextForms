using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NOTEBOOK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Form2 f2;
        private void fORM2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            f2= new Form2();
            f2.ShowDialog();
        }
    }
}
