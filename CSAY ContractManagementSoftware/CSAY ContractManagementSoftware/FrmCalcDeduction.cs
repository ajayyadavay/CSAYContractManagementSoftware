using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSAY_ContractManagementSoftware
{
    public partial class FrmCalcDeduction : Form
    {
        bool AddLastRow = true;

        public FrmCalcDeduction()
        {
            InitializeComponent();
        }

        private void FrmCalcDeduction_Load(object sender, EventArgs e)
        {
            //load deduction percentage format filename in combobox
            string dir = Environment.CurrentDirectory + "\\ComboBoxList\\DeductionPercent";
            string[] files = Directory.GetFiles(dir, "*.txt", SearchOption.AllDirectories);//Directory.GetFiles(dir);

            foreach (string filePath in files) ComboBoxFileFormat.Items.Add(System.IO.Path.GetFileName(filePath));

            string[] rownames = new string[] { "A", "B", "D", "H", "I", "L" };
            string[] colnames = new string[] { "Amount Estimate", "Amount Contract", "Amount Upto Previous", "Amount Upto This Bill", "Amount This Bill Only" };

            foreach (string rname in rownames) ComboBoxRow.Items.Add(rname);
            foreach (string cname in colnames) ComboBoxColumn.Items.Add(cname);
        }

        private void importDeductionPercentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            GenerateAmountDataGridFromText();
            AddLastRow = true;
        }

        private void GenerateAmountDataGridFromText()
        {
            string filename = Environment.CurrentDirectory + "\\ComboBoxList\\DeductionPercent\\" + ComboBoxFileFormat.Text;
            LoadTxtToDatagridview(dataGridView1, filename, 1, 3);
        }

        public void LoadTxtToDatagridview(DataGridView Dgv, string FileName, int TxtStartRow, int no_of_Col)
        {
            string[] ReadingText = new string[100];
            //string RWYCoordFilenName;
            int i;
            StreamReader sr;
            string line;


            line = "";
            //FileName = @".\InputFolder\" + TxtAirportCode.Text + "\\" + "Strip_RL.txt";
            //Pass the file path and file name to the StreamReader constructor
            sr = new StreamReader(FileName);
            //Read the first line of text
            line = sr.ReadLine();
            ReadingText[0] = line;
            //Continue to read until you reach end of file
            i = 1;
            while (line != null)
            {
                //Read the next line
                line = sr.ReadLine();
                ReadingText[i] = line;
                i++;
            }
            //close the file
            sr.Close();

            //load RL data of strip
            Dgv.Rows.Clear();
            int startrow = TxtStartRow;
            int sn = 1;
            for (int row = startrow; row < (i - startrow); row++)
            {
                Dgv.Rows.Add();
                //Dgv.Rows[row - startrow].Cells[0].Value = sn.ToString();
                sn++;
            }

            for (int row = startrow; row < (i - startrow); row++)
            {
                string[] splittedtext = ReadingText[row].Split('\t');
                for (int col = 0; col < no_of_Col; col++)
                {
                    Dgv.Rows[row - startrow].Cells[col].Value = splittedtext[col];
                }
            }

        }

        private void calculateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                int n = dataGridView1.RowCount - 1;
                if (AddLastRow == true)
                {
                    n = n;
                }
                else
                {
                    n = n - 1;
                }
                double subtotal = Convert.ToDouble(TxtAmount.Text);
                double percent, amount, sumpercent = 0, sumamount = 0;
                for (int i = 0; i < n; i++)
                {
                    percent = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    sumpercent += percent;
                    percent /= 100.0;
                    amount = Math.Round(subtotal * percent, 2);
                    sumamount += amount;
                    dataGridView1.Rows[i].Cells[3].Value = amount.ToString();
                    sumamount = Math.Round(sumamount, 2);
                    sumpercent = Math.Round(sumpercent, 2);
                }
                if (AddLastRow == true)
                {
                    dataGridView1.Rows.Add();
                    AddLastRow = false;

                }
                dataGridView1.Rows[n].Cells[1].Value = "Total";
                dataGridView1.Rows[n].Cells[2].Value = sumpercent.ToString();
                dataGridView1.Rows[n].Cells[3].Value = sumamount.ToString();


            }
            catch
            {

            }

        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void TxtGenerateRowNo_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int n = Convert.ToInt32(TxtGenerateRowNo.Text);
                dataGridView1.Rows.Clear();
                for (int i = 0; i < n; i++)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = (i + 1).ToString();
                }
                AddLastRow = true;
            }
            catch
            {

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnLoadAmount_Click(object sender, EventArgs e)
        {
            try
            {
                int rowno = 100, colno = 100, colno2 = 100;
                string str;

                string[] RowID = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L" };
                string[] ColID = new string[] {
                    "SN",
                    "Description Estimate",
                    "Amount Estimate", "Amount Contract", "Description Bill",
                    "Amount Upto Previous", "Amount Upto This Bill", "Amount This Bill Only"};

                for (int i = 0; i < 12; i++)
                {
                    str = RowID[i];
                    if (ComboBoxRow.Text == str)
                    {
                        rowno = i;
                    }
                }

                for (int i = 0; i < 8; i++)
                {
                    str = ColID[i];
                    if (ComboBoxColumn.Text == str)
                    {
                        colno = i;
                    }
                }

                if (colno <= 3) colno2 = 1;
                else if (colno >= 4 && colno < 8) colno2 = 4;

                if (rowno != 100 && colno != 100)
                {
                    FrmContract fc = (FrmContract)Application.OpenForms["FrmContract"];
                    TxtAmount.Text = (fc.dataGridView1.Rows[rowno].Cells[colno].Value).ToString();

                    LblAmount.Text = (fc.dataGridView1.Rows[rowno].Cells[colno2].Value).ToString() + "=";
                }
                else
                {
                    TxtAmount.Text = 0.ToString();
                    LblAmount.Text = "Amount =";
                }

            }
            catch
            {

            }
        }
    }
}
