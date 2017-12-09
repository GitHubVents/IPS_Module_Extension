using PDMWebService.Data.Solid.ElementsCase;
using System;
using System.Windows.Forms;
//using PDMWebService.Data.Solid.ElementsCase;


namespace Vents_PLM.Spigot
{
    public partial class SpigotControl : UserControl
    {
        double width, height;
        int spigotType;

         
        public SpigotControl()
        {
            InitializeComponent();
        }

        private void btnBuildSpigot_Click(object sender, EventArgs e)
        {
           SpigotBuilder spigot = new SpigotBuilder();

            if (ConvertValues())
            {
                //this.btnBuildSpigot.Click += spigot.CheckExistPart;

                spigot.Build(spigotType, new SolidWorksLibrary.Builders.ElementsCase.Vector2(width, height));
            }
        }

        private void comboBoxSpigotType_SelectedIndexChanged(object sender, EventArgs e)
        {            
            spigotType = Convert.ToInt32(comboBoxSpigotType.SelectedItem);
        }

        //private void InitializeComponent()
        //{
        //    this.SuspendLayout();
        //    // 
        //    // SpigotControl
        //    // 
        //    this.Name = "SpigotControl";
        //    this.Size = new System.Drawing.Size(627, 472);
        //    this.ResumeLayout(false);

        //}

        private bool ConvertValues()
        {
            try
            {
                width = Convert.ToDouble(txtBoxWidth.Text);
                height = Convert.ToDouble(txtBoxHeight.Text);
                return true;
            }
            catch (Exception)
            {
                MessageBox.Show("Введены некоректные данные");
                return false;
                throw;
            }
        }
    }
}