namespace CSAY_ContractManagementSoftware
{
    partial class FrmPBAmount
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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            dataGridViewPB = new DataGridView();
            ColDescription = new DataGridViewTextBoxColumn();
            ColPBPercent = new DataGridViewTextBoxColumn();
            ColTotalAmount = new DataGridViewTextBoxColumn();
            ColPBAmount = new DataGridViewTextBoxColumn();
            label1 = new Label();
            TxtCostEstimate = new TextBox();
            TxtContractPrice = new TextBox();
            label2 = new Label();
            TxtPercentbelow = new TextBox();
            label3 = new Label();
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            importAmountToolStripMenuItem = new ToolStripMenuItem();
            calculateToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripSeparator();
            exitToolStripMenuItem = new ToolStripMenuItem();
            RadioInclPS = new RadioButton();
            RadioExcludePS = new RadioButton();
            toolStripMenuItem2 = new ToolStripSeparator();
            ((System.ComponentModel.ISupportInitialize)dataGridViewPB).BeginInit();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridViewPB
            // 
            dataGridViewPB.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = Color.LightGray;
            dataGridViewPB.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridViewPB.BackgroundColor = Color.White;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = SystemColors.Control;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle2.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
            dataGridViewPB.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            dataGridViewPB.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewPB.Columns.AddRange(new DataGridViewColumn[] { ColDescription, ColPBPercent, ColTotalAmount, ColPBAmount });
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = SystemColors.Window;
            dataGridViewCellStyle3.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle3.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
            dataGridViewPB.DefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewPB.Location = new Point(30, 144);
            dataGridViewPB.Name = "dataGridViewPB";
            dataGridViewPB.RowHeadersVisible = false;
            dataGridViewPB.RowTemplate.Height = 25;
            dataGridViewPB.Size = new Size(618, 187);
            dataGridViewPB.TabIndex = 24;
            // 
            // ColDescription
            // 
            ColDescription.HeaderText = "Description";
            ColDescription.Name = "ColDescription";
            ColDescription.SortMode = DataGridViewColumnSortMode.NotSortable;
            ColDescription.Width = 200;
            // 
            // ColPBPercent
            // 
            ColPBPercent.HeaderText = "Percent";
            ColPBPercent.Name = "ColPBPercent";
            ColPBPercent.SortMode = DataGridViewColumnSortMode.NotSortable;
            // 
            // ColTotalAmount
            // 
            ColTotalAmount.HeaderText = "Total Amount to Calculate PB";
            ColTotalAmount.Name = "ColTotalAmount";
            ColTotalAmount.SortMode = DataGridViewColumnSortMode.NotSortable;
            ColTotalAmount.Width = 150;
            // 
            // ColPBAmount
            // 
            ColPBAmount.HeaderText = "Calculated PB Amount";
            ColPBAmount.Name = "ColPBAmount";
            ColPBAmount.SortMode = DataGridViewColumnSortMode.NotSortable;
            ColPBAmount.Width = 150;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label1.Location = new Point(30, 43);
            label1.Name = "label1";
            label1.Size = new Size(182, 20);
            label1.TabIndex = 25;
            label1.Text = "Cost Estimate without VAT";
            // 
            // TxtCostEstimate
            // 
            TxtCostEstimate.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            TxtCostEstimate.Location = new Point(218, 43);
            TxtCostEstimate.Name = "TxtCostEstimate";
            TxtCostEstimate.Size = new Size(159, 27);
            TxtCostEstimate.TabIndex = 26;
            TxtCostEstimate.TextChanged += TxtCostEstimate_TextChanged;
            // 
            // TxtContractPrice
            // 
            TxtContractPrice.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            TxtContractPrice.Location = new Point(218, 81);
            TxtContractPrice.Name = "TxtContractPrice";
            TxtContractPrice.Size = new Size(159, 27);
            TxtContractPrice.TabIndex = 28;
            TxtContractPrice.TextChanged += TxtContractPrice_TextChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label2.Location = new Point(30, 81);
            label2.Name = "label2";
            label2.Size = new Size(184, 20);
            label2.TabIndex = 27;
            label2.Text = "Contract Price without VAT";
            // 
            // TxtPercentbelow
            // 
            TxtPercentbelow.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            TxtPercentbelow.Location = new Point(272, 112);
            TxtPercentbelow.Name = "TxtPercentbelow";
            TxtPercentbelow.Size = new Size(105, 27);
            TxtPercentbelow.TabIndex = 30;
            TxtPercentbelow.Text = "85";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label3.Location = new Point(30, 115);
            label3.Name = "label3";
            label3.Size = new Size(222, 20);
            label3.TabIndex = 29;
            label3.Text = "Percent Below For Additional PB";
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(658, 28);
            menuStrip1.TabIndex = 31;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { importAmountToolStripMenuItem, toolStripMenuItem2, calculateToolStripMenuItem, toolStripMenuItem1, exitToolStripMenuItem });
            fileToolStripMenuItem.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(44, 24);
            fileToolStripMenuItem.Text = "File";
            // 
            // importAmountToolStripMenuItem
            // 
            importAmountToolStripMenuItem.Name = "importAmountToolStripMenuItem";
            importAmountToolStripMenuItem.Size = new Size(180, 24);
            importAmountToolStripMenuItem.Text = "Import Amount";
            importAmountToolStripMenuItem.Click += importAmountToolStripMenuItem_Click;
            // 
            // calculateToolStripMenuItem
            // 
            calculateToolStripMenuItem.Name = "calculateToolStripMenuItem";
            calculateToolStripMenuItem.Size = new Size(180, 24);
            calculateToolStripMenuItem.Text = "Calculate";
            calculateToolStripMenuItem.Click += calculateToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(177, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new Size(180, 24);
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
            // 
            // RadioInclPS
            // 
            RadioInclPS.AutoSize = true;
            RadioInclPS.Checked = true;
            RadioInclPS.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            RadioInclPS.Location = new Point(439, 58);
            RadioInclPS.Name = "RadioInclPS";
            RadioInclPS.Size = new Size(184, 24);
            RadioInclPS.TabIndex = 32;
            RadioInclPS.TabStop = true;
            RadioInclPS.Text = "Include Provisional Sum";
            RadioInclPS.UseVisualStyleBackColor = true;
            // 
            // RadioExcludePS
            // 
            RadioExcludePS.AutoSize = true;
            RadioExcludePS.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            RadioExcludePS.Location = new Point(439, 88);
            RadioExcludePS.Name = "RadioExcludePS";
            RadioExcludePS.Size = new Size(187, 24);
            RadioExcludePS.TabIndex = 33;
            RadioExcludePS.Text = "Exclude Provisional Sum";
            RadioExcludePS.UseVisualStyleBackColor = true;
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(177, 6);
            // 
            // FrmPBAmount
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(658, 343);
            Controls.Add(RadioExcludePS);
            Controls.Add(RadioInclPS);
            Controls.Add(TxtPercentbelow);
            Controls.Add(label3);
            Controls.Add(TxtContractPrice);
            Controls.Add(label2);
            Controls.Add(TxtCostEstimate);
            Controls.Add(label1);
            Controls.Add(dataGridViewPB);
            Controls.Add(menuStrip1);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MainMenuStrip = menuStrip1;
            MaximizeBox = false;
            Name = "FrmPBAmount";
            Text = "PB Amount";
            Load += FrmPBAmount_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridViewPB).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dataGridViewPB;
        private Label label1;
        private TextBox TxtCostEstimate;
        private TextBox TxtContractPrice;
        private Label label2;
        private TextBox TxtPercentbelow;
        private Label label3;
        private DataGridViewTextBoxColumn ColDescription;
        private DataGridViewTextBoxColumn ColPBPercent;
        private DataGridViewTextBoxColumn ColTotalAmount;
        private DataGridViewTextBoxColumn ColPBAmount;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem importAmountToolStripMenuItem;
        private ToolStripSeparator toolStripMenuItem1;
        private RadioButton RadioInclPS;
        private RadioButton RadioExcludePS;
        private ToolStripMenuItem calculateToolStripMenuItem;
        private ToolStripSeparator toolStripMenuItem2;
    }
}