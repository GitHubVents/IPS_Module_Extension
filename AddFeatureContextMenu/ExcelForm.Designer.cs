namespace AddFeatureContextMenu
{
    partial class ExcelForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.IMBASE_3ViewBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.button1.ForeColor = System.Drawing.Color.Black;
            this.button1.Location = new System.Drawing.Point(179, 62);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 39);
            this.button1.TabIndex = 2;
            this.button1.Text = "Добавить в IMBASE";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.button2.ForeColor = System.Drawing.Color.Black;
            this.button2.Location = new System.Drawing.Point(41, 62);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(105, 39);
            this.button2.TabIndex = 7;
            this.button2.Text = "Открыть файл Excel";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // IMBASE_3ViewBtn
            // 
            this.IMBASE_3ViewBtn.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.IMBASE_3ViewBtn.ForeColor = System.Drawing.Color.Black;
            this.IMBASE_3ViewBtn.Location = new System.Drawing.Point(0, 2);
            this.IMBASE_3ViewBtn.Name = "IMBASE_3ViewBtn";
            this.IMBASE_3ViewBtn.Size = new System.Drawing.Size(25, 25);
            this.IMBASE_3ViewBtn.TabIndex = 8;
            this.IMBASE_3ViewBtn.Text = "@";
            this.IMBASE_3ViewBtn.UseVisualStyleBackColor = false;
            this.IMBASE_3ViewBtn.Click += new System.EventHandler(this.IMBASE_3ViewBtn_Click);
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LimeGreen;
            this.ClientSize = new System.Drawing.Size(328, 142);
            this.Controls.Add(this.IMBASE_3ViewBtn);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "ExcelForm";
            this.Text = "ExcelForm";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button IMBASE_3ViewBtn;
    }
}