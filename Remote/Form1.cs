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

namespace Remote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        
        /// <summary>
        /// Будет выбирать и читать эксель файл
        /// </summary>
        /// <param name="sender">XLS файл</param>
        /// <param name="e">Таблица</param>       

        private void открытьФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var ReadExcel = new ExcelFile();
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK) //если файл вабран юзером
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo("xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = (DataTable)ReadExcel.ReadExcel(file.FileName);
                        DataGridXLS.Visible = true;
                        DataGridXLS.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Wrong format", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var SaveExcel = new ExcelFile();
            SaveFileDialog file = new SaveFileDialog();
            string fileExt = Path.GetExtension(file.FileName);
            SaveExcel.SaveExcel(fileExt);
            

        }
    }
}
