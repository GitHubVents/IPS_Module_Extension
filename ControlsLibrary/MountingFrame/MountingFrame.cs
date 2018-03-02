using PDMWebService.Data.Solid.ElementsCase;
using System;
using System.Windows.Forms;

namespace ControlsLibrary.MountingFrame
{
    public partial class MountingFrame : UserControl
    {
        int width;
        int length;
        int frameOffset;
        int frameType;
        int thickness;
        bool cheched;
        MountingFrameBuilder mountingFrContr;

        public MountingFrame()
        {
            InitializeComponent();
        }

        private void btnFrameBuild_Click(object sender, EventArgs e)
        {
            mountingFrContr = new MountingFrameBuilder();

            if (ConvertValues())
            {
                mountingFrContr.OpenDoc();
                mountingFrContr.BuildMountageFrame(cheched, width, length, thickness, frameType, frameOffset, null, null, false);
                mountingFrContr.SaveDoc();
            };
        }

        private bool ConvertValues()
        {
            try
            {
                width = Convert.ToInt32(textBoxWidth.Text);
                length = Convert.ToInt32(textBoxLenth.Text);
                if (frameType == 3)
                {
                    frameOffset = Convert.ToInt32(textBoxDisplacement.Text);
                }
                return true;
            }
            catch (Exception)
            {
                MessageBox.Show("Impossible to create the model. The input values are inappropriate.");
                return false;
            }
        }

        private void comboBoxFrameType_SelectedIndexChanged(object sender, EventArgs e)
        {
            frameType = comboBoxFrameType.SelectedIndex;

            switch (frameType)
            {
                case 0:
                    textBoxDisplacement.Enabled = false;
                    textBoxDisplacement.Text = "";
                    break;
                case 1:
                    textBoxDisplacement.Enabled = false;
                    textBoxDisplacement.Text = "";
                    break;
                case 2:
                    textBoxDisplacement.Enabled = false;
                    textBoxDisplacement.Text = "";
                    break;
                case 3:
                    textBoxDisplacement.Enabled = true;
                    break;
            }
        }

        private void comboBoxThickness_SelectedIndexChanged(object sender, EventArgs e)
        {
            thickness = Convert.ToInt32(comboBoxThickness.SelectedItem);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            cheched = checkBox1.Checked;
        }
    }
}