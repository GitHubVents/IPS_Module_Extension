using System;
using System.Data;
using System.Windows.Forms;

namespace AddFeatureContextMenu
{
    public partial class ExcelForm : Form
    {
        string pathToFile;
        
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
            //ips.ShowContextForm(1007);
            IPS_ExcelOperations pp = new IPS_ExcelOperations();
            pp.ShowContextForm(1053);

            pp.GetTags();
        }
    }
}