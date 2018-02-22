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
        IPS_ExcelOperations ips = new IPS_ExcelOperations();

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelFunctionality obj = new ExcelFunctionality(pathToFile); // вызвали конструктор класса, инициализация поля filePath

            DataSet ds = ExcelFunctionality.GetTotalSheetsData();
            ips.MainMethodForFillingImbaseTable(ds);
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

        private void IMBASE_3ViewBtn_Click(object sender, EventArgs e)
        {

            //List<ImbaseTableList> similarTables = new List<ImbaseTableList>() { new ImbaseTableList(12, "First"), new ImbaseTableList(44, "Second") };
            //int idNewType = ips.CreateTemporarType(similarTables);
            //  MessageBox.Show("idNewType  " + idNewType.ToString());
            ips.ShowContextForm(1007);
        }
    }
}