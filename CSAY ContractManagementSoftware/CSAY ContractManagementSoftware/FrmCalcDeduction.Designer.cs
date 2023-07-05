namespace CSAY_ContractManagementSoftware
{
    partial class FrmCalcDeduction
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
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            importDeductionPercentToolStripMenuItem = new ToolStripMenuItem();
            calculateToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripSeparator();
            clearToolStripMenuItem = new ToolStripMenuItem();
            exitToolStripMenuItem = new ToolStripMenuItem();
            dataGridView1 = new DataGridView();
            ColSN = new DataGridViewTextBoxColumn();
            ColDescription = new DataGridViewTextBoxColumn();
            ColPercentage = new DataGridViewTextBoxColumn();
            ColAmount = new DataGridViewTextBoxColumn();
            label1 = new Label();
            label2 = new Label();
            ComboBoxRow = new ComboBox();
            ComboBoxColumn = new ComboBox();
            TxtAmount = new TextBox();
            LblAmount = new Label();
            ComboBoxFileFormat = new ComboBox();
            label4 = new Label();
            label5 = new Label();
            TxtGenerateRowNo = new TextBox();
            BtnLoadAmount = new Button();
            menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(775, 28);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { importDeductionPercentToolStripMenuItem, calculateToolStripMenuItem, toolStripMenuItem1, clearToolStripMenuItem, exitToolStripMenuItem });
            fileToolStripMenuItem.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(44, 24);
            fileToolStripMenuItem.Text = "File";
            // 
            // importDeductionPercentToolStripMenuItem
            // 
            importDeductionPercentToolStripMenuItem.Name = "importDeductionPercentToolStripMenuItem";
            importDeductionPercentToolStripMenuItem.Size = new Size(248, 24);
            importDeductionPercentToolStripMenuItem.Text = "Import Deduction Percent";
            importDeductionPercentToolStripMenuItem.Click += importDeductionPercentToolStripMenuItem_Click;
            // 
            // calculateToolStripMenuItem
            // 
            calculateToolStripMenuItem.Name = "calculateToolStripMenuItem";
            calculateToolStripMenuItem.Size = new Size(248, 24);
            calculateToolStripMenuItem.Text = "Calculate";
            calculateToolStripMenuItem.Click += calculateToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(245, 6);
            // 
            // clearToolStripMenuItem
            // 
            clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            clearToolStripMenuItem.Size = new Size(248, 24);
            clearToolStripMenuItem.Text = "Clear";
            clearToolStripMenuItem.Click += clearToolStripMenuItem_Click;
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new Size(248, 24);
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle1.BackColor = Color.FromArgb(224, 224, 224);
            dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView1.BackgroundColor = Color.White;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = SystemColors.Control;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle2.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
            dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { ColSN, ColDescription, ColPercentage, ColAmount });
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = SystemColors.Window;
            dataGridViewCellStyle3.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle3.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
            dataGridView1.DefaultCellStyle = dataGridViewCellStyle3;
            dataGridView1.Location = new Point(12, 158);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(755, 218);
            dataGridView1.TabIndex = 1;
            // 
            // ColSN
            // 
            ColSN.HeaderText = "SN";
            ColSN.Name = "ColSN";
            ColSN.SortMode = DataGridViewColumnSortMode.NotSortable;
            // 
            // ColDescription
            // 
            ColDescription.HeaderText = "Description";
            ColDescription.Name = "ColDescription";
            ColDescription.SortMode = DataGridViewColumnSortMode.NotSortable;
            ColDescription.Width = 300;
            // 
            // ColPercentage
            // 
            ColPercentage.HeaderText = "Percentage";
            ColPercentage.Name = "ColPercentage";
            ColPercentage.SortMode = DataGridViewColumnSortMode.NotSortable;
            // 
            // ColAmount
            // 
            ColAmount.HeaderText = "Amount";
            ColAmount.Name = "ColAmount";
            ColAmount.SortMode = DataGridViewColumnSortMode.NotSortable;
            ColAmount.Width = 200;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label1.Location = new Point(14, 90);
            label1.Name = "label1";
            label1.Size = new Size(135, 20);
            label1.TabIndex = 2;
            label1.Text = "Choose Row Name";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label2.Location = new Point(155, 90);
            label2.Name = "label2";
            label2.Size = new Size(157, 20);
            label2.TabIndex = 3;
            label2.Text = "Choose Column Name";
            // 
            // ComboBoxRow
            // 
            ComboBoxRow.DropDownStyle = ComboBoxStyle.DropDownList;
            ComboBoxRow.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            ComboBoxRow.FormattingEnabled = true;
            ComboBoxRow.Location = new Point(14, 115);
            ComboBoxRow.Name = "ComboBoxRow";
            ComboBoxRow.Size = new Size(135, 28);
            ComboBoxRow.TabIndex = 4;
            // 
            // ComboBoxColumn
            // 
            ComboBoxColumn.DropDownStyle = ComboBoxStyle.DropDownList;
            ComboBoxColumn.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            ComboBoxColumn.FormattingEnabled = true;
            ComboBoxColumn.Location = new Point(155, 115);
            ComboBoxColumn.Name = "ComboBoxColumn";
            ComboBoxColumn.Size = new Size(332, 28);
            ComboBoxColumn.TabIndex = 5;
            // 
            // TxtAmount
            // 
            TxtAmount.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            TxtAmount.Location = new Point(536, 116);
            TxtAmount.Name = "TxtAmount";
            TxtAmount.Size = new Size(227, 27);
            TxtAmount.TabIndex = 6;
            // 
            // LblAmount
            // 
            LblAmount.AutoSize = true;
            LblAmount.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            LblAmount.Location = new Point(536, 92);
            LblAmount.Name = "LblAmount";
            LblAmount.Size = new Size(76, 20);
            LblAmount.TabIndex = 7;
            LblAmount.Text = "Amount =";
            // 
            // ComboBoxFileFormat
            // 
            ComboBoxFileFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            ComboBoxFileFormat.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            ComboBoxFileFormat.FormattingEnabled = true;
            ComboBoxFileFormat.Location = new Point(194, 48);
            ComboBoxFileFormat.Name = "ComboBoxFileFormat";
            ComboBoxFileFormat.Size = new Size(336, 28);
            ComboBoxFileFormat.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label4.Location = new Point(12, 51);
            label4.Name = "label4";
            label4.Size = new Size(176, 20);
            label4.TabIndex = 8;
            label4.Text = "Choose .txt File to Import";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label5.Location = new Point(536, 51);
            label5.Name = "label5";
            label5.Size = new Size(153, 20);
            label5.TabIndex = 11;
            label5.Text = "Generate No. of Rows";
            // 
            // TxtGenerateRowNo
            // 
            TxtGenerateRowNo.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            TxtGenerateRowNo.Location = new Point(692, 48);
            TxtGenerateRowNo.Name = "TxtGenerateRowNo";
            TxtGenerateRowNo.Size = new Size(71, 27);
            TxtGenerateRowNo.TabIndex = 10;
            TxtGenerateRowNo.TextChanged += TxtGenerateRowNo_TextChanged;
            // 
            // BtnLoadAmount
            // 
            BtnLoadAmount.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            BtnLoadAmount.Location = new Point(493, 115);
            BtnLoadAmount.Name = "BtnLoadAmount";
            BtnLoadAmount.Size = new Size(37, 29);
            BtnLoadAmount.TabIndex = 12;
            BtnLoadAmount.Text = ">>";
            BtnLoadAmount.UseVisualStyleBackColor = true;
            BtnLoadAmount.Click += BtnLoadAmount_Click;
            // 
            // FrmCalcDeduction
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(775, 386);
            Controls.Add(BtnLoadAmount);
            Controls.Add(label5);
            Controls.Add(TxtGenerateRowNo);
            Controls.Add(ComboBoxFileFormat);
            Controls.Add(label4);
            Controls.Add(LblAmount);
            Controls.Add(TxtAmount);
            Controls.Add(ComboBoxColumn);
            Controls.Add(ComboBoxRow);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(dataGridView1);
            Controls.Add(menuStrip1);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MainMenuStrip = menuStrip1;
            MaximizeBox = false;
            Name = "FrmCalcDeduction";
            Text = "FrmCalcDeduction";
            Load += FrmCalcDeduction_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem importDeductionPercentToolStripMenuItem;
        private ToolStripMenuItem calculateToolStripMenuItem;
        private ToolStripSeparator toolStripMenuItem1;
        private ToolStripMenuItem clearToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private DataGridView dataGridView1;
        private DataGridViewTextBoxColumn ColSN;
        private DataGridViewTextBoxColumn ColDescription;
        private DataGridViewTextBoxColumn ColPercentage;
        private DataGridViewTextBoxColumn ColAmount;
        private Label label1;
        private Label label2;
        private ComboBox ComboBoxRow;
        private ComboBox ComboBoxColumn;
        private TextBox TxtAmount;
        private Label LblAmount;
        private ComboBox ComboBoxFileFormat;
        private Label label4;
        private Label label5;
        private TextBox TxtGenerateRowNo;
        private Button BtnLoadAmount;
    }
}