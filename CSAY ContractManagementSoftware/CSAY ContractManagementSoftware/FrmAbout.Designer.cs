namespace CSAY_ContractManagementSoftware
{
    partial class FrmAbout
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAbout));
            groupBox1 = new GroupBox();
            textBox1 = new TextBox();
            groupBox2 = new GroupBox();
            textBox2 = new TextBox();
            BtnExit = new Button();
            groupBox3 = new GroupBox();
            textBox3 = new TextBox();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(textBox1);
            groupBox1.Font = new Font("Comic Sans MS", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox1.ForeColor = Color.DodgerBlue;
            groupBox1.Location = new Point(21, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(711, 258);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Creator";
            // 
            // textBox1
            // 
            textBox1.ForeColor = SystemColors.MenuHighlight;
            textBox1.Location = new Point(18, 27);
            textBox1.Multiline = true;
            textBox1.Name = "textBox1";
            textBox1.ReadOnly = true;
            textBox1.Size = new Size(664, 211);
            textBox1.TabIndex = 1;
            textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(textBox2);
            groupBox2.Font = new Font("Comic Sans MS", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox2.ForeColor = Color.OrangeRed;
            groupBox2.Location = new Point(21, 278);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(711, 228);
            groupBox2.TabIndex = 2;
            groupBox2.TabStop = false;
            groupBox2.Text = "Bill Calculation Formula";
            // 
            // textBox2
            // 
            textBox2.ForeColor = SystemColors.MenuHighlight;
            textBox2.Location = new Point(18, 27);
            textBox2.Multiline = true;
            textBox2.Name = "textBox2";
            textBox2.ReadOnly = true;
            textBox2.Size = new Size(664, 186);
            textBox2.TabIndex = 1;
            textBox2.TabStop = false;
            textBox2.Text = resources.GetString("textBox2.Text");
            // 
            // BtnExit
            // 
            BtnExit.FlatAppearance.BorderColor = Color.FromArgb(255, 128, 0);
            BtnExit.FlatAppearance.MouseDownBackColor = Color.FromArgb(255, 128, 128);
            BtnExit.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 224, 192);
            BtnExit.FlatStyle = FlatStyle.Flat;
            BtnExit.Font = new Font("Comic Sans MS", 11F, FontStyle.Regular, GraphicsUnit.Point);
            BtnExit.ForeColor = Color.Black;
            BtnExit.Location = new Point(613, 628);
            BtnExit.Name = "BtnExit";
            BtnExit.Size = new Size(119, 35);
            BtnExit.TabIndex = 3;
            BtnExit.Text = "Exit";
            BtnExit.UseVisualStyleBackColor = true;
            BtnExit.Click += BtnExit_Click;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(textBox3);
            groupBox3.Font = new Font("Comic Sans MS", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox3.ForeColor = Color.DarkViolet;
            groupBox3.ImeMode = ImeMode.NoControl;
            groupBox3.Location = new Point(21, 512);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(711, 110);
            groupBox3.TabIndex = 3;
            groupBox3.TabStop = false;
            groupBox3.Text = "Filter Format";
            // 
            // textBox3
            // 
            textBox3.ForeColor = SystemColors.MenuHighlight;
            textBox3.Location = new Point(18, 27);
            textBox3.Multiline = true;
            textBox3.Name = "textBox3";
            textBox3.ReadOnly = true;
            textBox3.Size = new Size(664, 75);
            textBox3.TabIndex = 1;
            textBox3.Text = "You can add as many filter as required by putting keywork AND, value should in single quote and column name without quote.\r\nColumn Name1 = 'Value1' AND Column Name2 = 'Value2'\r\n";
            // 
            // FrmAbout
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(744, 687);
            Controls.Add(groupBox3);
            Controls.Add(BtnExit);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Name = "FrmAbout";
            Text = "About";
            Load += FrmAbout_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox1;
        private TextBox textBox1;
        private GroupBox groupBox2;
        private TextBox textBox2;
        private Button BtnExit;
        private GroupBox groupBox3;
        private TextBox textBox3;
    }
}