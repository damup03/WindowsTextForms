using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelDataReader;

using Aspose.Cells;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Xls;

namespace WFAPr
{
    public partial class Form1 : Form
    {
        private string filename = string.Empty;

        private DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();

        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;

                    Text = filename;

                    OpenExcelFile(filename);
                }
                else
                {
                    throw new Exception("Вы ничего не выбрали");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ОШИБКА!", MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }

        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = db.Tables;

            toolStripComboBox1.Items.Clear();

            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }

            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection
                [Convert.ToString(toolStripComboBox1.SelectedItem)];

            dataGridView1.DataSource = table;
        }

        private void экспортWordToolStripMenuItem_Click(object sender, EventArgs e) // Экспорт
        {
            
            //Серьезность	Код	Описание	Проект	Файл	Строка	Состояние подавления
            //Ошибка CS1061  "Form1" не содержит определения "экспортWordToolStripMenuItem_Click",
            //и не удалось найти доступный метод расширения "экспортWordToolStripMenuItem_Click",
            //принимающий тип "Form1" в качестве первого аргумента(возможно, пропущена директива using или ссылка на сборку).	
            //WFAPr C:\Users\user\OneDrive\Рабочий стол\WFAPr\WFAPr\Form1.Designer.cs  92  Активные

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

        Form2 f2;
        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            f2 = new Form2();
            f2.Show();
        }

        private void вWorldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.doc)|*.doc";

            sfd.FileName = "export.doc";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                ToCsV(dataGridView1, sfd.FileName);

            }
        }

        private void вExcellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Excel Documents (*.xls)|*.xls";

            sfd.FileName = "export.xls";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                ToCsV(dataGridView1, sfd.FileName);

            }
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("Excel.xls");

            DocxSaveOptions options = new DocxSaveOptions();
            options.ClearData = true;
            options.CreateDirectory = true;
            options.CachedFileFolder = "cache";
            options.MergeAreas = true;

            workbook.Save(@"D:\Users\PRexportfileexcel1.docx", options);
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
