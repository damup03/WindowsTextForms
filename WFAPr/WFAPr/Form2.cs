using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WFAPr
{
    public partial class Form2 : Form
    {
        List<string> names = new List<string>();
        List<double[]> data = new List<double[]>();

        public Form2()
        {
            InitializeComponent();

            names.Add("Курс");
            data.Add(new double[]
        {
       1,
       2,
       3
        });

            names.Add("Группа");
            data.Add(new double[]
        {
       4110,
       4111,
       4112
        });
            names.Add("Направление /специальность");
            data.Add(new double[]
        {

        });
            names.Add("Семестр");
            data.Add(new double[]
        {

        });
            names.Add("Вид практики");
            data.Add(new double[]
        {

        });
            names.Add("Дата проведения собрания");
            data.Add(new double[]
        {

        });
            dataGridView1.DataSource = GetResultsTable();
        }

        public DataTable GetResultsTable()
        {
            DataTable d = new DataTable();

            for (int i = 0; i < this.data.Count; i++)
            {
                string name = this.names[i];

                d.Columns.Add(name);

                List<object> objectNumbers = new List<object>();

                foreach (double number in this.data[i])
                {
                    objectNumbers.Add((object)number);
                }

                while (d.Rows.Count < objectNumbers.Count)
                {
                    d.Rows.Add();
                }

                for (int a = 0; a < objectNumbers.Count; a++)
                {
                    d.Rows[a][i] = objectNumbers[a];
                }
            }
            return d;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void вWordToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.doc)|*.doc";

            sfd.FileName = "export.doc";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                ToCsV(dataGridView1, sfd.FileName);

            }
        }

        private void вExcellToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Excel Documents (*.xls)|*.xls";

            sfd.FileName = "export.xls";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                ToCsV(dataGridView1, sfd.FileName);

            }
        }

        private void ToCsV(DataGridView dGV, string filename)
        {

            string stOutput = "";

            string sHeaders = "";

            for (int j = 0; j < dGV.Columns.Count; j++)

                sHeaders = sHeaders.ToString() + Convert.ToString(dGV.Columns[j].HeaderText) + "\t";

            stOutput += sHeaders + "\r\n";

            for (int i = 0; i < dGV.RowCount - 1; i++)
            {

                string stLine = "";

                for (int j = 0; j < dGV.Rows[i].Cells.Count; j++)

                    stLine = stLine.ToString() + Convert.ToString(dGV.Rows[i].Cells[j].Value) + "\t";

                stOutput += stLine + "\r\n";

            }

            Encoding utf16 = Encoding.GetEncoding(1254);

            byte[] output = utf16.GetBytes(stOutput);

            FileStream fs = new FileStream(filename, FileMode.Create);

            BinaryWriter bw = new BinaryWriter(fs);

            bw.Write(output, 0, output.Length);

            bw.Flush();

            bw.Close();

            fs.Close();

        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
