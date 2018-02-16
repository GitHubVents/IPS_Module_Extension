using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddFeatureContextMenu
{
    public partial class ExcelForm : Form
    {
        string pathToFile;
        string name;
        int count = 0;
        bool wholeFile = false;
        
        public ExcelForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (pathToFile != string.Empty && pathToFile != null)
            {
                ExcelFunctionality obj = new ExcelFunctionality(pathToFile); // вызвали конструктор класса, инициализация поля filePath
                obj.MainMethodForFillingImbaseTable();

                MessageBox.Show("Success!");
            }
            
            //dataGridExcel.DataSource = ExcelFunctionality.GetTotalSheetsData().Tables[count];
        }


        private void ConvertValues()
        {
            try
            {
                //name = textName.Text;
                count = Convert.ToInt32( textCount.Text);
                wholeFile = checkBox1.Checked;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Введены недопустимые значения");
                throw ex;
            }
        }


        private string GetExcelPath()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "All Files (*.*)|*.*";
            dialog.Multiselect = false;
            DialogResult result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                return System.IO.Path.GetFullPath(dialog.FileName);
            }
            return string.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pathToFile = GetExcelPath();
        }
    }
}