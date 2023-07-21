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
    public partial class FrmPBAmount : Form
    {
        public FrmPBAmount()
        {
            InitializeComponent();
        }

        private void FrmPBAmount_Load(object sender, EventArgs e)
        {
            GeneratePBAmountDataGridFromText();
        }

        private void ImportAmount()
        {
            double tempPS, tempST, total;
            FrmContract fc = (FrmContract)Application.OpenForms["FrmContract"];
            tempPS = Convert.ToDouble(fc.dataGridView1.Rows[0].Cells[2].Value);
            tempST = Convert.ToDouble(fc.dataGridView1.Rows[1].Cells[2].Value);
            if (RadioInclPS.Checked == true) total = tempPS + tempST;
            else total = tempST;
            TxtCostEstimate.Text = total.ToString();

            tempPS = Convert.ToDouble(fc.dataGridView1.Rows[0].Cells[3].Value);
            tempST = Convert.ToDouble(fc.dataGridView1.Rows[1].Cells[3].Value);
            if (RadioInclPS.Checked == true) total = tempPS + tempST;
            else total = tempST;
            TxtContractPrice.Text = total.ToString();
        }

        private void EstimateAmountToTable()
        {
            dataGridViewPB.Rows[0].Cells[2].Value = TxtContractPrice.Text;
            dataGridViewPB.Rows[3].Cells[2].Value = TxtContractPrice.Text;

            double est, cont, temp85, per;
            est = Convert.ToDouble(TxtCostEstimate.Text);
            cont = Convert.ToDouble(TxtContractPrice.Text);
            per = Convert.ToDouble(TxtPercentbelow.Text);
            per /= 100.0;
            if (cont > est * per)
            {
                temp85 = 0.0;
            }
            else
            {
                temp85 = Math.Round(est * per - cont, 2);
            }
            dataGridViewPB.Rows[1].Cells[2].Value = temp85.ToString();

        }

        private void GeneratePBAmountDataGridFromText()
        {
            string filename = Environment.CurrentDirectory + "\\ComboBoxList\\" + "PBAmountCalc.txt";
            LoadTxtToDatagridview(dataGridViewPB, filename, 1, 2);
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

        private void TxtCostEstimate_TextChanged(object sender, EventArgs e)
        {
            try
            {
                EstimateAmountToTable();
            }
            catch
            {

            }
        }

        private void TxtContractPrice_TextChanged(object sender, EventArgs e)
        {
            try
            {
                EstimateAmountToTable();
            }
            catch
            {

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void importAmountToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ImportAmount();
            }
            catch
            {

            }
        }

        private void calculateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                double num1, per, sum, total = 0;
                for (int i = 0; i <= 3; i++)
                {
                    if (i != 2)
                    {
                        num1 = Convert.ToDouble(dataGridViewPB.Rows[i].Cells[2].Value);
                        per = Convert.ToDouble(dataGridViewPB.Rows[i].Cells[1].Value);
                        sum = Math.Round(num1 * per / 100.0, 2);
                        dataGridViewPB.Rows[i].Cells[3].Value = sum.ToString();
                        total += sum;
                    }
                    else
                    {
                        total = Math.Round(total, 2);
                        dataGridViewPB.Rows[i].Cells[3].Value = total.ToString();
                    }
                }

            }
            catch
            {

            }
        }
    }
}
