namespace Vents_PLM.Spigot
{
    partial class SpigotControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxSpigotType = new System.Windows.Forms.ComboBox();
            this.txtBoxWidth = new System.Windows.Forms.TextBox();
            this.txtBoxHeight = new System.Windows.Forms.TextBox();
            this.btnBuildSpigot = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(32, 61);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(174, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Вибровставка";
            // 
            // comboBoxSpigotType
            // 
            this.comboBoxSpigotType.FormattingEnabled = true;
            this.comboBoxSpigotType.Items.AddRange(new object[] {
            "20",
            "30"});
            this.comboBoxSpigotType.Location = new System.Drawing.Point(37, 160);
            this.comboBoxSpigotType.Name = "comboBoxSpigotType";
            this.comboBoxSpigotType.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSpigotType.TabIndex = 1;
            this.comboBoxSpigotType.SelectedIndexChanged += new System.EventHandler(this.comboBoxSpigotType_SelectedIndexChanged);
            // 
            // txtBoxWidth
            // 
            this.txtBoxWidth.Location = new System.Drawing.Point(200, 160);
            this.txtBoxWidth.Name = "txtBoxWidth";
            this.txtBoxWidth.Size = new System.Drawing.Size(100, 20);
            this.txtBoxWidth.TabIndex = 2;
            // 
            // txtBoxHeight
            // 
            this.txtBoxHeight.Location = new System.Drawing.Point(366, 161);
            this.txtBoxHeight.Name = "txtBoxHeight";
            this.txtBoxHeight.Size = new System.Drawing.Size(100, 20);
            this.txtBoxHeight.TabIndex = 3;
            // 
            // btnBuildSpigot
            // 
            this.btnBuildSpigot.Location = new System.Drawing.Point(542, 157);
            this.btnBuildSpigot.Name = "btnBuildSpigot";
            this.btnBuildSpigot.Size = new System.Drawing.Size(129, 23);
            this.btnBuildSpigot.TabIndex = 4;
            this.btnBuildSpigot.Text = "Построить";
            this.btnBuildSpigot.UseVisualStyleBackColor = true;
            this.btnBuildSpigot.Click += new System.EventHandler(this.btnBuildSpigot_Click);
            // 
            // SpigotControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.btnBuildSpigot);
            this.Controls.Add(this.txtBoxHeight);
            this.Controls.Add(this.txtBoxWidth);
            this.Controls.Add(this.comboBoxSpigotType);
            this.Controls.Add(this.label1);
            this.Name = "SpigotControl";
            this.Size = new System.Drawing.Size(828, 572);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxSpigotType;
        private System.Windows.Forms.TextBox txtBoxWidth;
        private System.Windows.Forms.TextBox txtBoxHeight;
        private System.Windows.Forms.Button btnBuildSpigot;
    }
}
