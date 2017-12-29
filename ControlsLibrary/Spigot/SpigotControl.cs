using System;
using System.Windows.Forms;
using PDMWebService.Data.Solid.ElementsCase;

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

        public void Build(object sender, EventArgs e)
        {
            SpigotBuilder spigot = new SpigotBuilder();
            if (ConvertValues())
            {
                spigot.Build(spigotType, new SolidWorksLibrary.Builders.ElementsCase.Vector2(width, height));// метод из библиотеки SolidWorksLibrary
            }
        }

        private void comboBoxSpigotType_SelectedIndexChanged(object sender, EventArgs e)
        {            
            spigotType = Convert.ToInt32(comboBoxSpigotType.SelectedItem);
        }
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