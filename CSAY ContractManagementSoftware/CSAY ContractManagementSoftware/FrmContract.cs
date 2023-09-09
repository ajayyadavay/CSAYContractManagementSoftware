using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SQLite;
using iText.Kernel.Pdf;
//using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.IO.Font.Constants;
using iText.Kernel.Pdf.Canvas.Draw;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;
using NodaTime;
using AD_BS_Converter;

namespace CSAY_ContractManagementSoftware
{
    public partial class FrmContract : Form
    {
#nullable disable
        //Declare global variables
        int rowsAmtGrid = 12, colsAmtGrid = 8;
        int rowsAmtGridBill = 12;
        int n_of_calc_Row = 4, n_of_calc_Col = 5;

        string Cur_Dir, Contract_ID, Ward, Project_Type, Project_Folders, ThisContractFolder, EventHistoryFolder, LastEventFolder;
        string FYFolder;
        public FrmContract()
        {
            InitializeComponent();
        }

        private void FrmContract_Load(object sender, EventArgs e)
        {
            //Control1.TabPages.Remove(TabLetter);
            //string tdate = DateTime.UtcNow.ToString("MM-dd-yyyy");
            string tdate = DateTime.UtcNow.ToString("yyyy-MM-dd");
            TxtToday.Text = tdate;


            //load format bill in combobox
            string dir = Environment.CurrentDirectory + "\\ComboBoxList\\BillFormat";
            string[] files = Directory.GetFiles(dir, "*.txt", SearchOption.AllDirectories);//Directory.GetFiles(dir);

            foreach (string filePath in files) ComboBoxFormatBill.Items.Add(System.IO.Path.GetFileName(filePath));


            //load default format bill format text name
            string[] FormatbillList = System.IO.File.ReadAllLines(@".\ComboBoxList\DefaultFormatBill.txt");
            TxtFormatBill.Text = FormatbillList[0];
            /*foreach (var line in FormatbillList)
            {
                ComboBoxFY.Items.Add(line);
            }*/

            GenerateAmountDataGridFromText();
            //GenerateAmountDataGrid();


            SetGridColorAndFont();
            SetColorofInputCells();


            //Add ---> Procurement category 
            string[] ProcCategoryList = System.IO.File.ReadAllLines(@".\ComboBoxList\ProcurementCategory.txt");
            foreach (var line in ProcCategoryList)
            {
                ComboBoxProCategory.Items.Add(line);
            }



            //Add ---> Fiscal year 
            string[] FiscalYearList = System.IO.File.ReadAllLines(@".\ComboBoxList\FiscalYear.txt");
            foreach (var line in FiscalYearList)
            {
                ComboBoxFY.Items.Add(line);
            }

            //Add ---> Budget Type 
            string[] BudgetTypeList = System.IO.File.ReadAllLines(@".\ComboBoxList\BudgetType.txt");
            foreach (var line in BudgetTypeList)
            {
                ComboBoxBudgetType.Items.Add(line);
            }

            //Add ---> CurrentStatus
            string[] CurrentStatusList = System.IO.File.ReadAllLines(@".\ComboBoxList\CurrentStatus.txt");
            foreach (var line in CurrentStatusList)
            {
                ComboBoxCurrentStatus.Items.Add(line);
            }

            //Add ---> Project Type
            string[] ProjectTypeList = System.IO.File.ReadAllLines(@".\ComboBoxList\ProjectType.txt");
            foreach (var line in ProjectTypeList)
            {
                ComboBoxProjectType.Items.Add(line);
            }

            //Add ---> Filter
            string[] FilterList = System.IO.File.ReadAllLines(@".\ComboBoxList\Filter.txt");
            foreach (var line in FilterList)
            {
                ComboBoxFilterBy1.Items.Add(line);
            }

            //Add ---> APG1BankName
            string[] APG1BankNameList = System.IO.File.ReadAllLines(@".\ComboBoxList\BankName.txt");
            foreach (var line in APG1BankNameList)
            {
                ComboBoxAGP1BankName.Items.Add(line);
                ComboBoxAGP2BankName.Items.Add(line);
                ComboBoxPBBankName.Items.Add(line);
            }

            //Add ---> InsBankName
            string[] InsBankNameList = System.IO.File.ReadAllLines(@".\ComboBoxList\InsuranceName.txt");
            foreach (var line in InsBankNameList)
            {
                ComboBoxInsBankName.Items.Add(line);
            }

            //Add ---> Public Entity Name
            string[] PEList = System.IO.File.ReadAllLines(@".\ComboBoxList\PublicEntity.txt");
            foreach (var line in PEList)
            {
                ComboBoxPE.Items.Add(line);
            }

            GenerateUnicodeDateTable();
            GenerateAPGTippaniTable();

        }

        private void GenerateUnicodeDateTable()
        {
            dataGridView3.Rows.Clear();
            dataGridView3.Rows.Add();
            dataGridView3.Rows[0].Cells[0].Value = "Contract Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[1].Cells[0].Value = "Work Permit Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[2].Cells[0].Value = "New Work Complete Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[3].Cells[0].Value = "Old Work Complete Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[4].Cells[0].Value = "EOT letter Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[5].Cells[0].Value = "APG1 Issue date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[6].Cells[0].Value = "APG2 Issue Date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[7].Cells[0].Value = "APG1 Deadline date";
            dataGridView3.Rows.Add();
            dataGridView3.Rows[8].Cells[0].Value = "APG2 Deadline Date";
        }

        private void GenerateAPGTippaniTable()
        {
            dataGridView4.Rows.Clear();
            dataGridView4.Rows.Add();
            dataGridView4.Rows[0].Cells[0].Value = "Letter requesting AP date (BS)";
            dataGridView4.Rows.Add();
            dataGridView4.Rows[1].Cells[0].Value = "Percentage of AP";
            dataGridView4.Rows.Add();
            dataGridView4.Rows[2].Cells[0].Value = "Tippani writing date (BS)";
            dataGridView4.Rows.Add();
            dataGridView4.Rows[3].Cells[0].Value = "APG Amount";
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
            //MessageBox.Show("i, startrow, (i-startrow) = " + i.ToString() + ", " + startrow.ToString() + ", " + (i - startrow).ToString());
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


        private void GenerateAmountDataGridFromText()
        {
            string filename = Environment.CurrentDirectory + "\\ComboBoxList\\BillFormat\\" + TxtFormatBill.Text;
            LoadTxtToDatagridview(dataGridView1, filename, 1, 8);
        }

        public void GenerateAmountDataGrid() //Function to generate Amount Data grid
        {
            //initialize and declared variables
            string[] DescriptionEstimate = new string[]
            {
                "PS", "Subtotal", "VAT %", "VAT Amount", "Contingency %", "Physical Contingency %",
                "Price Contingency", "Total (A+B+D)", "GrandTotal incl. contingencies", "Advance Payment1",
                "Advance Payment2", "Total Advance Payment"
            };

            string[] DescriptionBill = new string[]
            {
                "PS", "Subtotal", "VAT %", "VAT Amount", "Contingency %", "Physical Contingency %",
                "Price Contingency %", "Total (A+B+D)", "GrandTotal incl. contingencies", "Advance Payment deduct %",
                "Deduct Retention %", "Net Payment to Contractor"
            };

            string[] SNEstimate = new string[]
            {
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"
            };

            int[] percentages = new int[]
            {
                13, //VAT %
                3, //Contingency %
                0, //Physical contingency %
                0  //Price contingency %
            };
            int[] ColIndex = new int[]
            {
                2, //Estimate Column
                3, //Contract Column
                //5, //Amount up2 previous column
                6,  //Amount up2 this bill column
                //7  //Amount of This bill only column
            };

            int[] RowIndex = new int[]
            {
                2, //VAT % Row
                4, //Contingency % Row
                5, //Physical contingency % Row
                6,  //Price contingency % Row
            };

            //generate rows in contract and bill datagrid
            for (int i = 0; i < rowsAmtGrid; i++) //0 to 11
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = SNEstimate[i]; //SN of Estimate
                dataGridView1.Rows[i].Cells[1].Value = DescriptionEstimate[i];//Description of Estimate
                dataGridView1.Rows[i].Cells[4].Value = DescriptionBill[i]; // Description of Bill
            }

            //Entering % data like VAT, Contingencies, etc.
            for (int i = 0; i < n_of_calc_Col - 2; i++) //Column Index 5no.
            {
                for (int j = 0; j < n_of_calc_Row; j++) //Row Index and Percentage data index 4no.
                {
                    dataGridView1.Rows[RowIndex[j]].Cells[ColIndex[i]].Value = percentages[j].ToString("0.000"); //Write all percentage in respective cells
                }
            }

            dataGridView1.Rows[9].Cells[2].Value = 0.ToString("0.000"); //AP1 of column Estimate %
            dataGridView1.Rows[10].Cells[2].Value = 0.ToString("0.000"); //AP2 of column Estimate %

            dataGridView1.Rows[10].Cells[6].Value = 5.ToString("0.000"); //Retention Deduciton %
            dataGridView1.Rows[10].Cells[7].Value = 5.ToString("0.000"); //Retention Deduciton %

            //making AmountUp2Previous column all values zero
            for (int j = 0; j < rowsAmtGrid; j++) //j = 0 to 11
            {
                dataGridView1.Rows[j].Cells[5].Value = 0.ToString("0.000"); //Write all percentage in respective cells
            }
            dataGridView1.Rows[10].Cells[5].Value = 5.ToString("0.000"); //Retention Deduciton %
        }

        public void SetGridColorAndFont()
        {
            /*dataGridView1.DefaultCellStyle.Font = new Font("Comic Sans MS", 12);
            dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.White;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Purple;*/

        }

        private void ComboBoxFY_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtFY.Text = ComboBoxFY.Text;
        }

        private void ComboBoxBudgetType_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtBudgetType.Text = ComboBoxBudgetType.Text;
        }

        private void ComboBoxProjectType_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtProjectType.Text = ComboBoxProjectType.Text;
        }

        private void Fun_AnalyseDate()
        {
            try
            {
                int rem_days = DifferenceInDate(TxtToday.Text, TxtWorkComplete.Text);
                TxtDaysRem.Text = rem_days.ToString();
                //int ContractDays = DifferenceInDate(TxtWorkPermit.Text, TxtWorkComplete.Text);
                int ContractDays = DifferenceInDate(TxtContractAgreement.Text, TxtWorkComplete.Text);

                //days remaining from today
                if (rem_days > 0)
                {
                    TxtWorkComplete.BackColor = Color.LightGreen;
                    TxtDateAnalysis.ForeColor = Color.ForestGreen;
                    TxtDaysRem.ForeColor = Color.ForestGreen;
                    TxtDateAnalysis.Text = "OK." + rem_days + " days remaining for completion.";
                }
                else if (rem_days <= 0)
                {
                    TxtWorkComplete.BackColor = Color.LightCoral;
                    TxtDateAnalysis.ForeColor = Color.Red;
                    TxtDaysRem.ForeColor = Color.Red;
                    TxtDateAnalysis.Text = "REVIEW. " + rem_days + " days past Deadline.";
                }

                //determine Minimum APG Date,Insurance, PB date
                int APGDay = Convert.ToInt32(ContractDays * 0.8 + 1);
                TxtAPG1MinDL.Text = NewDateAFterAddingDays_and_Months(APGDay, 0, TxtContractAgreement.Text);//txtworkpermit.text
                TxtAPG2MinDL.Text = TxtAPG1MinDL.Text;
                TxtInsMinDL.Text = TxtWorkComplete.Text;
                TxtPBMinDL.Text = NewDateAFterAddingDays_and_Months(365, 1, TxtWorkComplete.Text);
                TxtFL_PBMinDL.Text = TxtWorkComplete.Text;

                //check if APG, PB, Ins document deadline is equal or more than Min valid date
                int tempdays;
                //APG1
                tempdays = DifferenceInDate(TxtAPG1MinDL.Text, TxtAPG1DL.Text);
                if (tempdays >= 0)
                {
                    TxtAPG1Remark.Text = "Valid";
                    TxtAPG1Remark.ForeColor = Color.ForestGreen;
                }
                else if (tempdays < 0)
                {
                    TxtAPG1Remark.Text = "Review";
                    TxtAPG1Remark.ForeColor = Color.Red;
                }
                //APG2
                tempdays = DifferenceInDate(TxtAPG2MinDL.Text, TxtAPG2DL.Text);
                if (tempdays >= 0)
                {
                    TxtAPG2Remark.Text = "Valid";
                    TxtAPG2Remark.ForeColor = Color.ForestGreen;
                }
                else if (tempdays < 0)
                {
                    TxtAPG2Remark.Text = "Review";
                    TxtAPG2Remark.ForeColor = Color.Red;
                }
                //PB
                tempdays = DifferenceInDate(TxtPBMinDL.Text, TxtPBDL.Text);
                if (tempdays >= 0)
                {
                    TxtPBRemark.Text = "Valid";
                    TxtPBRemark.ForeColor = Color.ForestGreen;
                }
                else if (tempdays < 0)
                {
                    TxtPBRemark.Text = "Review";
                    TxtPBRemark.ForeColor = Color.Red;
                }
                //Insurance
                tempdays = DifferenceInDate(TxtInsMinDL.Text, TxtInsDL.Text);
                if (tempdays >= 0)
                {
                    TxtInsRemark.Text = "Valid";
                    TxtInsRemark.ForeColor = Color.ForestGreen;
                }
                else if (tempdays < 0)
                {
                    TxtInsRemark.Text = "Review";
                    TxtInsRemark.ForeColor = Color.Red;
                }
                //PB_FL
                tempdays = DifferenceInDate(TxtFL_PBMinDL.Text, TxtFL_PBDeadline.Text);
                if (tempdays >= 0)
                {
                    TxtFL_PBRemark.Text = "Valid";
                    TxtFL_PBRemark.ForeColor = Color.ForestGreen;
                }
                else if (tempdays < 0)
                {
                    TxtFL_PBRemark.Text = "Review";
                    TxtFL_PBRemark.ForeColor = Color.Red;
                }

                //checking APG,PB,Ins date from Today
                //APG1
                tempdays = DifferenceInDate(TxtToday.Text, TxtAPG1DL.Text);
                if (tempdays > 7)
                {
                    TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.Red;
                }
                //APG2
                tempdays = DifferenceInDate(TxtToday.Text, TxtAPG2DL.Text);
                if (tempdays > 7)
                {
                    TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.Red;
                }
                //PB
                tempdays = DifferenceInDate(TxtToday.Text, TxtPBDL.Text);
                if (tempdays > 7)
                {
                    TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.Red;
                }
                //Insurance
                tempdays = DifferenceInDate(TxtToday.Text, TxtInsDL.Text);
                if (tempdays > 7)
                {
                    TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.Red;
                }

                //PB_FL
                tempdays = DifferenceInDate(TxtToday.Text, TxtFL_PBDeadline.Text);
                if (tempdays > 7)
                {
                    TxtPBFLDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    TxtPBFLDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    TxtPBFLDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.Red;
                }
            }
            catch
            {

            }
        }

        private int DifferenceInDate(string StartDate, string EndDate)
        {
            try
            {
                //Date should be in format YYYY-MM-DD e.g. 2022-02-23
                int year1, month1, days1, year2, month2, days2;
                string[] temp_date1 = StartDate.Split("-");
                string[] temp_date2 = EndDate.Split("-");

                /*int[] monthdays = new int[]
                {
                    31,28,31,30,31,30,31,31,30,31,30,31
                };*/

                year1 = Convert.ToInt32(temp_date1[0]);
                month1 = Convert.ToInt32(temp_date1[1]);
                days1 = Convert.ToInt32(temp_date1[2]);

                year2 = Convert.ToInt32(temp_date2[0]);
                month2 = Convert.ToInt32(temp_date2[1]);
                days2 = Convert.ToInt32(temp_date2[2]);

                DateTime start = new DateTime(year1, month1, days1);
                DateTime end = new DateTime(year2, month2, days2);

                TimeSpan difference = end - start; //create TimeSpan object

                return difference.Days;


            }
            catch
            {
                return 0;
            }

        }

        private string NewDateAFterAddingDays_and_Months(int DaysToAdd, int MonthsToAdd, string OldDate)
        {
            try
            {
                //Date should be in format YYYY-MM-DD e.g. 2022-02-23
                int year1, month1, days1, year2, month2, days2;
                string[] temp_date1 = OldDate.Split("-");

                year1 = Convert.ToInt32(temp_date1[0]);
                month1 = Convert.ToInt32(temp_date1[1]);
                days1 = Convert.ToInt32(temp_date1[2]);

                DateTime start = new DateTime(year1, month1, days1);
                DateTime somedate = start.AddDays(DaysToAdd);
                somedate = somedate.AddMonths(MonthsToAdd);

                year2 = somedate.Year;
                month2 = somedate.Month;
                days2 = somedate.Day;

                OldDate = year2 + "-" + month2 + "-" + days2;

                return OldDate;
            }
            catch
            {
                return "";
            }
        }

        private void DeleteTextFields()
        {
            TxtFY.Text = "";
            TxtContractID.Text = "";
            TxtContractName.Text = "";
            TxtContractBudget.Text = "";
            TxtWard.Text = "";
            TxtProjectType.Text = "";
            TxtBudgetType.Text = "";
            TxtLocation.Text = "";

            TxtAPG1RefNo.Text = "";
            TxtAPG1DL.Text = "";
            TxtAPG1Amount.Text = "";
            TxtAPG1MinDL.Text = "";
            TxtAPG1Remark.Text = "";

            TxtAPG2RefNo.Text = "";
            TxtAPG2DL.Text = "";
            TxtAPG2Amount.Text = "";
            TxtAPG2MinDL.Text = "";
            TxtAPG2Remark.Text = "";

            TxtPBRefNo.Text = "";
            TxtPBDL.Text = "";
            TxtPBAmount.Text = "";
            TxtPBMinDL.Text = "";
            TxtPBRemark.Text = "";

            TxtInsRefNo.Text = "";
            TxtInsDL.Text = "";
            TxtInsAmount.Text = "";
            TxtInsMinDL.Text = "";
            TxtInsRemark.Text = "";

            TxtCurrentStatus.Text = "";
            TxtNoticeIssued.Text = "";
            TxtLOI.Text = "";
            TxtLOA.Text = "";
            TxtContractAgreement.Text = "";
            TxtWorkPermit.Text = "";
            TxtWorkComplete.Text = "";
            TxtRunningBill.Text = "";
            TxtFinalBill.Text = "";
            TxtDaysRem.Text = "";

            TxtContractorName.Text = "";
            TxtAddressOfContractor.Text = "";
            TxtEmail1.Text = "";
            TxtContractorOther.Text = "";

            TxtProjectDescription.Text = "";
            TxtLength.Text = "";
            TxtBreadth.Text = "";
            TxtHeight.Text = "";

            TxtContractorNameDev.Text = "";
            TxtContractorAddressDev.Text = "";

            TxtDateAnalysis.Text = "Log";
            TxtAPG1DaysRem.Text = "";
            TxtAPG2DaysRem.Text = "";
            TxtPBDaysRem.Text = "";
            TxtInsDaysRem.Text = "";

            TxtBankNameAPG1.Text = "";
            TxtBankNameAPG2.Text = "";
            TxtBankNamePB.Text = "";
            TxtBankNameIns.Text = "";

            TxtBankAddressAPG1.Text = "";
            TxtBankAddressAPG2.Text = "";
            TxtBankAddressPB.Text = "";
            TxtBankAddressIns.Text = "";

            TxtProcurementcategory.Text = "";
            TxtProcurementMethod.Text = "";
            TxtTotalEstimatedAmount.Text = "";
            TxtTotalContractAmount.Text = "";
            TxtTotalFinalBillAmount.Text = "";
            //TxtPE.Text = "";

            TxtFL_PBRef.Text = "";
            TxtFL_PBDeadline.Text = "";
            TxtFL_PBAmount.Text = "";
            TxtFL_PBMinDL.Text = "";
            TxtFL_PBRemark.Text = "";
            TxtFL_BankNamePB.Text = "";
            TxtFL_BankAddressPB.Text = "";
            TxtPBFLDaysRem.Text = "";

            ComboBoxFY.SelectedIndex = -1;
            ComboBoxBudgetType.SelectedIndex = -1;
            ComboBoxProjectType.SelectedIndex = -1;
            ComboBoxCurrentStatus.SelectedIndex = -1;
            ComboBoxAGP1BankName.SelectedIndex = -1;
            ComboBoxAGP2BankName.SelectedIndex = -1;
            ComboBoxPBBankName.SelectedIndex = -1;
            ComboBoxInsBankName.SelectedIndex = -1;
        }

        private bool IsContractIDUnique()
        {
            bool C_ID = false;

            string value;
            SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
            ConnectDb.Open();

            //for unique value
            string query = "SELECT DISTINCT " + "ContractID" + " FROM ContractTable";
            SQLiteDataAdapter DataAdptr = new SQLiteDataAdapter(query, ConnectDb);

            DataTable Dt = new DataTable();
            DataAdptr.Fill(Dt);

            //ComboBoxDistinctVal1.Items.Clear();
            string thisID = TxtContractID.Text;
            foreach (DataRow row in Dt.Rows)
            {
                value = row[0].ToString();
                if (thisID == value)
                {
                    C_ID = true;
                    break;
                }
                else
                {
                    C_ID = false;
                }
                //ComboBoxDistinctVal1.Items.Add(value);
            }
            ConnectDb.Close();

            return C_ID;
        }

        private void Fun_Add(object sender, EventArgs e)
        {
            string FiscalYear = TxtFY.Text;
            string ContractID = TxtContractID.Text;
            string ContractName = TxtContractName.Text;
            string ContractBudget = TxtContractBudget.Text;
            string Ward = TxtWard.Text;
            string ProjectType = TxtProjectType.Text;
            string BudgetType = TxtBudgetType.Text;
            string Location = TxtLocation.Text;

            string APG1DocRefNo = TxtAPG1RefNo.Text;
            string APG1Deadline = TxtAPG1DL.Text;
            string APG1Amount = TxtAPG1Amount.Text;
            string APG1MinDL = TxtAPG1MinDL.Text;
            string APG1Remark = TxtAPG1Remark.Text;

            string APG2DocRefNo = TxtAPG2RefNo.Text;
            string APG2Deadline = TxtAPG2DL.Text;
            string APG2Amount = TxtAPG2Amount.Text;
            string APG2MinDL = TxtAPG2MinDL.Text;
            string APG2Remark = TxtAPG2Remark.Text;

            string PBDocRefNo = TxtPBRefNo.Text;
            string PBDeadline = TxtPBDL.Text;
            string PBAmount = TxtPBAmount.Text;
            string PBMinDL = TxtPBMinDL.Text;
            string PBRemark = TxtPBRemark.Text;

            string InsDocRefNo = TxtInsRefNo.Text;
            string InsDeadline = TxtInsDL.Text;
            string InsAmount = TxtInsAmount.Text;
            string InsMinDL = TxtInsMinDL.Text;
            string InsRemark = TxtInsRemark.Text;

            string CurrentStatus = TxtCurrentStatus.Text;
            string NoticeIssued = TxtNoticeIssued.Text;
            string LOI = TxtLOI.Text;
            string LOA = TxtLOA.Text;
            string ContractAgreement = TxtContractAgreement.Text;
            string WorkPermit = TxtWorkPermit.Text;
            string WorkComplete = TxtWorkComplete.Text;
            string RunningBill = TxtRunningBill.Text;
            string FinalBill = TxtFinalBill.Text;
            string DaysRemaining = TxtDaysRem.Text;

            string NameOfContractor = TxtContractorName.Text;
            string AddressOfContractor = TxtAddressOfContractor.Text;
            string Email1 = TxtEmail1.Text;
            string ContractorOther = TxtContractorOther.Text;

            string ProjectDescription = TxtProjectDescription.Text;
            string Length = TxtLength.Text;
            string Breadth = TxtBreadth.Text;
            string Height = TxtHeight.Text;

            string ContractorNameDev = TxtContractorNameDev.Text;
            string ContractorAddressDev = TxtContractorAddressDev.Text;

            string APG1DaysRem = TxtAPG1DaysRem.Text;
            string APG2DaysRem = TxtAPG2DaysRem.Text;
            string PBDaysRem = TxtPBDaysRem.Text;
            string InsDaysRem = TxtInsDaysRem.Text;

            string APG1BankName = TxtBankNameAPG1.Text;
            string APG2BankName = TxtBankNameAPG2.Text;
            string PBBankName = TxtBankNamePB.Text;
            string InsBankName = TxtBankNameIns.Text;

            string APG1BankAddress = TxtBankAddressAPG1.Text;
            string APG2BankAddress = TxtBankAddressAPG2.Text;
            string PBBankAddress = TxtBankAddressPB.Text;
            string InsBankAddress = TxtBankAddressIns.Text;

            string ProcurementCategory = TxtProcurementcategory.Text;
            string ProcurementMethod = TxtProcurementMethod.Text;
            string TotalEstimatedAmount = TxtTotalEstimatedAmount.Text;
            string TotalContractAmount = TxtTotalContractAmount.Text;
            string TotalFinalBillAmount = TxtTotalFinalBillAmount.Text;
            string PublicEntity = TxtPE.Text;

            string PB2DocRefNo = TxtFL_PBRef.Text;
            string PB2Deadline = TxtFL_PBDeadline.Text;
            string PB2Amount = TxtFL_PBAmount.Text;
            string PB2MinDL = TxtFL_PBMinDL.Text;
            string PB2Remark = TxtFL_PBRemark.Text;
            string PB2BankName = TxtFL_BankNamePB.Text;
            string PB2BankAddress = TxtFL_BankAddressPB.Text;
            string PB2DaysRem = TxtPBFLDaysRem.Text;


            if (TxtFY.Text == "" || TxtContractID.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "")
            {
                TxtLog.Text += "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                TxtLog.Text += Environment.NewLine;
            }
            else
            {
                DialogResult dr = MessageBox.Show("Are you sure, you want to Add all data to Database?", "Add", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    //Add
                    SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
                    ConnectDb.Open();
                    string query = "INSERT INTO ContractTable(FiscalYear,ContractID,ContractName,ContractBudget,Ward," +
                        "ProjectType,BudgetType,Location,APG1DocRefNo,APG1Deadline, APG1Amount,APG1MinDL,APG1Remark," +
                        "APG2DocRefNo,APG2Deadline, APG2Amount,APG2MinDL,APG2Remark," +
                        "PBDocRefNo,PBDeadline, PBAmount,PBMinDL,PBRemark," +
                        "InsDocRefNo,InsDeadline, InsAmount,InsMinDL,InsRemark," +
                        "CurrentStatus,NoticeIssued,LOI,LOA,ContractAgreement,WorkPermit,WorkComplete,RunningBill,FinalBill,DaysRemaining," +
                        "NameOfContractor,AddressOfContractor,Email1,ContractorOther,ProjectDescription,Length,Breadth,Height,ContractorNameDev,ContractorAddressDev," +
                        "APG1DaysRem,APG2DaysRem,PBDaysRem,InsDaysRem,APG1BankName,APG2BankName ,PBBankName ,InsBankName,APG1BankAddress,APG2BankAddress,PBBankAddress,InsBankAddress," +
                        "ProcurementCategory, ProcurementMethod, TotalEstimatedAmount, TotalContractAmount, TotalFinalBillAmount, PublicEntity," +
                        "PB2DocRefNo,PB2Deadline, PB2Amount,PB2MinDL,PB2Remark,PB2BankName,PB2BankAddress,PB2DaysRem) " +
                        "VALUES('" + FiscalYear + "','" + ContractID + "','" + ContractName + "','" + ContractBudget + "'," +
                        "'" + Ward + "','" + ProjectType + "','" + BudgetType + "','" + Location + "'" +
                        ",'" + APG1DocRefNo + "','" + APG1Deadline + "','" + APG1Amount + "','" + APG1MinDL + "','" + APG1Remark + "'" +
                        ",'" + APG2DocRefNo + "','" + APG2Deadline + "','" + APG2Amount + "','" + APG2MinDL + "','" + APG2Remark + "'" +
                        ",'" + PBDocRefNo + "','" + PBDeadline + "','" + PBAmount + "','" + PBMinDL + "','" + PBRemark + "'" +
                        ",'" + InsDocRefNo + "','" + InsDeadline + "','" + InsAmount + "','" + InsMinDL + "','" + InsRemark + "'" +
                        ",'" + CurrentStatus + "','" + NoticeIssued + "','" + LOI + "','" + LOA + "','" + ContractAgreement + "','" + WorkPermit + "'" +
                        ",'" + WorkComplete + "','" + RunningBill + "','" + FinalBill + "','" + DaysRemaining + "'" +
                        ",'" + NameOfContractor + "','" + AddressOfContractor + "','" + Email1 + "','" + ContractorOther + "'" +
                        ",'" + ProjectDescription + "','" + Length + "','" + Breadth + "','" + Height + "','" + ContractorNameDev + "','" + ContractorAddressDev + "'" +
                        ",'" + APG1DaysRem + "','" + APG2DaysRem + "','" + PBDaysRem + "','" + InsDaysRem + "'" +
                        ",'" + APG1BankName + "','" + APG2BankName + "','" + PBBankName + "','" + InsBankName + "'" +
                        ",'" + APG1BankAddress + "','" + APG2BankAddress + "','" + PBBankAddress + "','" + InsBankAddress + "'" +
                        ",'" + ProcurementCategory + "','" + ProcurementMethod + "','" + TotalEstimatedAmount + "', '" + TotalContractAmount + "', '" + TotalFinalBillAmount + "', '" + PublicEntity + "' " +
                        ",'" + PB2DocRefNo + "','" + PB2Deadline + "','" + PB2Amount + "','" + PB2MinDL + "','" + PB2Remark + "','" + PB2BankName + "','" + PB2BankAddress + "','" + PB2DaysRem + "' )";// one data format  = '" + Height + "'

                    SQLiteCommand Cmd = new SQLiteCommand(query, ConnectDb);
                    Cmd.ExecuteNonQuery();

                    ConnectDb.Close();

                    //BtnCreateProjectFolder_Click(sender, e);
                    Fun_CreateProjectFolder();
                    BtnSave2Txt_Click(sender, e);
                    BtnResetBill_Click(sender, e);


                    // clear text boxes
                    TxtProjectID.Text = "";
                    DeleteTextFields();

                    TxtDateAnalysis.Text = "Log";
                    TxtAPG1DaysRem.Text = "";
                    TxtAPG2DaysRem.Text = "";
                    TxtPBDaysRem.Text = "";
                    TxtInsDaysRem.Text = "";
                    TxtPBFLDaysRem.Text = "";

                    ComboBoxFY.SelectedIndex = -1;
                    ComboBoxBudgetType.SelectedIndex = -1;
                    ComboBoxProjectType.SelectedIndex = -1;
                    ComboBoxCurrentStatus.SelectedIndex = -1;
                    ComboBoxProCategory.SelectedIndex = -1;
                    ComboBoxProMethod.SelectedIndex = -1;

                    Initial_State_of_Label();

                    TxtLog.AppendText("Activity: Record Successfully Added : " + ContractID + " of " + Ward + " at " + Location);
                    TxtLog.AppendText(Environment.NewLine);

                    /*using (System.IO.StreamWriter sw = System.IO.File.AppendText(@".\Log\Log.txt"))
                    {
                        Text2Write = "[" + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "]" + "  --->  " + "ADD" + " ---> " + ProjectName + " of " + Ward + " at " + Location; ;
                        sw.WriteLine(Text2Write);
                    }*/
                }
                else if (dr == DialogResult.No)
                {
                    //Nothing to do
                }
            }
        }

        private void Initial_State_of_Label()
        {
            TxtWorkComplete.BackColor = Color.White;
            TxtDateAnalysis.ForeColor = Color.Black;
            TxtDaysRem.ForeColor = Color.Black;

            TxtAPG1Remark.ForeColor = Color.Black;
            TxtAPG2Remark.ForeColor = Color.Black;
            TxtPBRemark.ForeColor = Color.Black;
            TxtInsRemark.ForeColor = Color.Black;

            TxtAPG1DaysRem.ForeColor = Color.Black;
            TxtAPG2DaysRem.ForeColor = Color.Black;
            TxtPBDaysRem.ForeColor = Color.Black;
            TxtInsDaysRem.ForeColor = Color.Black;
            TxtPBFLDaysRem.ForeColor = Color.Black;
        }

        private void BtnLoadAllRecord_Click(object sender, EventArgs e)
        {
            SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
            ConnectDb.Open();

            string query = "SELECT * FROM ContractTable";
            SQLiteDataAdapter DataAdptr = new SQLiteDataAdapter(query, ConnectDb);

            DataTable Dt = new DataTable();
            DataAdptr.Fill(Dt);
            dataGridView2.DataSource = Dt;

            ConnectDb.Close();
            LblDbLog.Text = "Recent Activity: Contract Record Loaded Successfully";

            int rcount = Dt.Rows.Count;
            LblRecordNo.Text = "Total No. of Record loaded:  " + rcount.ToString();
        }


        private void BtnToday_Click(object sender, EventArgs e)
        {
            string tdate = DateTime.UtcNow.ToString("yyyy-MM-dd");
            TxtToday.Text = tdate;
        }

        private void ComboBoxFilterBy1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Blue;
            RichTxtFilter.SelectedText += ComboBoxFilterBy1.Text;

            string value;
            SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
            ConnectDb.Open();

            //for unique value
            string query = "SELECT DISTINCT " + ComboBoxFilterBy1.Text + " FROM ContractTable";
            SQLiteDataAdapter DataAdptr = new SQLiteDataAdapter(query, ConnectDb);

            DataTable Dt = new DataTable();
            DataAdptr.Fill(Dt);

            ComboBoxDistinctVal1.Items.Clear();
            foreach (DataRow row in Dt.Rows)
            {
                value = row[0].ToString();
                ComboBoxDistinctVal1.Items.Add(value);
            }

            ConnectDb.Close();
        }

        private void ComboBoxDistinctVal1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Black;
            RichTxtFilter.SelectedText += "'" + ComboBoxDistinctVal1.Text + "'";
        }

        private void BtnFilter_Click(object sender, EventArgs e)
        {
            SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
            ConnectDb.Open();

            string query = "SELECT * FROM ContractTable where " + RichTxtFilter.Text;

            SQLiteDataAdapter DataAdptr = new SQLiteDataAdapter(query, ConnectDb);

            DataTable Dt = new DataTable();
            DataAdptr.Fill(Dt);
            dataGridView2.DataSource = Dt;


            ConnectDb.Close();
            //MessageBox.Show("Parameters Data Loaded Successfully.", "Load Parameters");
            int rcount = Dt.Rows.Count;
            LblRecordNo.Text = "Total No. of Record loaded:  " + rcount.ToString();
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            RichTxtFilter.Text = "";
        }

        private void BtnAND_Click(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Green;
            RichTxtFilter.SelectedText += " AND ";
        }

        private void BtnOR_Click(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Green;
            RichTxtFilter.SelectedText += " OR ";
        }

        private void BtnEqualTo_Click(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "=";
        }

        private void BtnLessThan_Click(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "<";

        }

        private void BtnGreaterThan_Click(object sender, EventArgs e)
        {
            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += ">";
        }

        private void ComboBoxAGP1BankName_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtBankNameAPG1.Text = ComboBoxAGP1BankName.Text;
        }

        private void ComboBoxAGP2BankName_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtBankNameAPG2.Text = ComboBoxAGP2BankName.Text;
        }

        private void ComboBoxPBBankName_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtBankNamePB.Text = ComboBoxPBBankName.Text;
        }

        private void ComboBoxInsBankName_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtBankNameIns.Text = ComboBoxInsBankName.Text;
        }

        private void BtnGuaranteeDate_Click(object sender, EventArgs e)
        {
            RichTxtFilter.Text = "";

            RichTxtFilter.SelectionColor = Color.Blue;
            RichTxtFilter.SelectedText += "APG1DaysRem";

            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "<=";

            RichTxtFilter.SelectionColor = Color.Black;
            RichTxtFilter.SelectedText += "'" + "7" + "'";

            RichTxtFilter.SelectionColor = Color.Green;
            RichTxtFilter.SelectedText += " OR ";

            RichTxtFilter.SelectionColor = Color.Blue;
            RichTxtFilter.SelectedText += "APG2DaysRem";

            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "<=";

            RichTxtFilter.SelectionColor = Color.Black;
            RichTxtFilter.SelectedText += "'" + "7" + "'";

            RichTxtFilter.SelectionColor = Color.Green;
            RichTxtFilter.SelectedText += " OR ";

            RichTxtFilter.SelectionColor = Color.Blue;
            RichTxtFilter.SelectedText += "PBDaysRem";

            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "<=";

            RichTxtFilter.SelectionColor = Color.Black;
            RichTxtFilter.SelectedText += "'" + "7" + "'";

            RichTxtFilter.SelectionColor = Color.Green;
            RichTxtFilter.SelectedText += " OR ";

            RichTxtFilter.SelectionColor = Color.Blue;
            RichTxtFilter.SelectedText += "InsDaysRem";

            RichTxtFilter.SelectionColor = Color.Red;
            RichTxtFilter.SelectedText += "<=";

            RichTxtFilter.SelectionColor = Color.Black;
            RichTxtFilter.SelectedText += "'" + "7" + "'";

        }

        private void Fun_CreateAllPdf()
        {
            //BillFileName = EventHistoryFolder + "\\Bill.txt";
            string ThisDir = Environment.CurrentDirectory;
            //string FontDir1 = ThisDir + "\\Font\\Preeti Normal.otf";
            // path folder
            CreateAccessProjectFolders();
            string PdfFileName = EventHistoryFolder + "\\PdfReport.pdf";
            //PdfWriter writer = new PdfWriter("E:\\AllPdf.pdf");
            PdfWriter writer = new PdfWriter(PdfFileName);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            //PdfFont KalimatiFont = PdfFontFactory.CreateFont(FontDir0, PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);
            //PdfFont PreetiFont = PdfFontFactory.CreateFont(FontDir1, PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);

            Paragraph header = new Paragraph();
            header.Add(TxtPE.Text + "\n" + "Record of Contract ID: " + TxtContractID.Text + " at " + TxtCurrentStatus.Text)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(14);
            //.SetFont(KalimatiFont);
            document.Add(header);

            Paragraph generated = new Paragraph();
            //generated.Add("Report Generated on : " + DateTime.UtcNow.ToString("yyyy-MM-dd|HH : mm : ss"))
            generated.Add("Report Generated on : " + DateTime.Now.ToString("F"))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetFontSize(10);
            //.SetFont(KalimatiFont);
            document.Add(generated);

            /*Paragraph generated1 = new Paragraph();
            generated1.Add("Report Generated from: CSAY CivilOne Software")
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetFontSize(12);
            //.SetFont(KalimatiFont);
            document.Add(generated1);*/

            //Line separator
            LineSeparator ls = new LineSeparator(new SolidLine());
            document.Add(ls);

            Paragraph generated2 = new Paragraph();
            generated2.Add("\n");
            //.SetTextAlignment(TextAlignment.RIGHT)
            //.SetFontSize(12);
            //.SetFont(KalimatiFont);
            document.Add(generated2);

            // Table
            iText.Layout.Element.Table table = new iText.Layout.Element.Table(3, false);

            //Row0------------------------------------------------------
            Cell cell00 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("A. General"));

            //Row1------------------------------------------------------
            Cell cell11 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("1"));

            Cell cell12 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Project ID"));
            Cell cell13 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtProjectID.Text));


            //Row2------------------------------------------------------
            Cell cell21 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("2"));

            Cell cell22 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Fiscal Year"));
            Cell cell23 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtFY.Text));

            //Row3------------------------------------------------------
            Cell cell31 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("3"));

            Cell cell32 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract ID"));
            Cell cell33 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractID.Text));

            //Row4------------------------------------------------------
            Cell cell41 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("4"));

            Cell cell42 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract Name"));
            Cell cell43 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractName.Text));

            //Row5------------------------------------------------------
            Cell cell51 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("5"));

            Cell cell52 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract Budget"));
            Cell cell53 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractBudget.Text));

            //Row6------------------------------------------------------
            Cell cell61 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("6"));

            Cell cell62 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ward and Location"));
            Cell cell63 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPE.Text + " - " + TxtWard.Text + ", " + TxtLocation.Text));

            //Row7------------------------------------------------------
            Cell cell71 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("7"));

            Cell cell72 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract Budget Type"));
            Cell cell73 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtBudgetType.Text));

            //Row8------------------------------------------------------
            Cell cell81 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("8"));

            Cell cell82 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract Project Type"));
            Cell cell83 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtProjectType.Text));

            //Row01------------------------------------------------------
            Cell cell01 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("B. Event Dates AD (YYYY-MM-DD)"));

            //Row9------------------------------------------------------
            Cell cell91 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("9"));

            Cell cell92 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Current Status"));
            Cell cell93 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtCurrentStatus.Text));

            //Row10------------------------------------------------------
            Cell cell101 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("10"));

            Cell cell102 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Notice Issued"));
            Cell cell103 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtNoticeIssued.Text));

            //Row11------------------------------------------------------
            Cell cell111 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("11"));

            Cell cell112 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Lettr of Intent (LOI)"));
            Cell cell113 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtLOI.Text));

            //Row12------------------------------------------------------
            Cell cell121 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("12"));

            Cell cell122 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Lettr of Acceptance (LOA)"));
            Cell cell123 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtLOA.Text));

            //Row13------------------------------------------------------
            Cell cell131 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("13"));

            Cell cell132 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contract Agreement"));
            Cell cell133 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractAgreement.Text));

            //Row14------------------------------------------------------
            Cell cell141 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("14"));

            Cell cell142 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Work Permit"));
            Cell cell143 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtWorkPermit.Text));

            //Row15------------------------------------------------------
            Cell cell151 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("15"));

            Cell cell152 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Work Completion"));
            Cell cell153 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtWorkComplete.Text));

            //Row16------------------------------------------------------
            Cell cell161 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("16"));

            Cell cell162 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Last Running Bill"));
            Cell cell163 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtRunningBill.Text));

            //Row17------------------------------------------------------
            Cell cell171 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("17"));

            Cell cell172 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Final Bill"));
            Cell cell173 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtFinalBill.Text));

            //Row18------------------------------------------------------
            Cell cell181 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("18"));

            Cell cell182 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Completion days remaining from Today"));
            Cell cell183 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtDaysRem.Text));

            //Row02------------------------------------------------------
            Cell cell02 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("C. Advance Payment 1 Bank Guarantee"));

            //Row19------------------------------------------------------
            Cell cell191 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("19"));

            Cell cell192 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Document Reference no."));
            Cell cell193 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1RefNo.Text));

            //Row20------------------------------------------------------
            Cell cell201 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("20"));

            Cell cell202 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Deadline"));
            Cell cell203 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1DL.Text));


            //Row21------------------------------------------------------
            Cell cell211 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("21"));

            Cell cell212 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Amount"));
            Cell cell213 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1Amount.Text));

            //Row22------------------------------------------------------
            Cell cell221 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("22"));

            Cell cell222 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Minimum DeadLine"));
            Cell cell223 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1MinDL.Text));

            //Row23------------------------------------------------------
            Cell cell231 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("23"));

            Cell cell232 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Remark"));
            Cell cell233 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1Remark.Text));

            //Row24------------------------------------------------------
            Cell cell241 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("24"));

            Cell cell242 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 Bank Name and Address"));
            Cell cell243 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               //.Add(new Paragraph(TxtBankNameAPG1.Text + ", " + TxtBankAddressAPG1.Text).SetFont(KalimatiFont));
               .Add(new Paragraph(TxtBankNameAPG1.Text + ", " + TxtBankAddressAPG1.Text));


            //Row03------------------------------------------------------
            Cell cell03 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("D. Advance Payment 2 Bank Guarantee"));

            //Row25------------------------------------------------------
            Cell cell251 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("25"));

            Cell cell252 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Document Reference no."));
            Cell cell253 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2RefNo.Text));

            //Row26------------------------------------------------------
            Cell cell261 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("26"));

            Cell cell262 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Deadline"));
            Cell cell263 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2DL.Text));


            //Row27------------------------------------------------------
            Cell cell271 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("27"));

            Cell cell272 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Amount"));
            Cell cell273 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2Amount.Text));

            //Row28------------------------------------------------------
            Cell cell281 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("28"));

            Cell cell282 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Minimum DeadLine"));
            Cell cell283 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2MinDL.Text));

            //Row29------------------------------------------------------
            Cell cell291 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("29"));

            Cell cell292 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Remark"));
            Cell cell293 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2Remark.Text));

            //Row30------------------------------------------------------
            Cell cell301 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("30"));

            Cell cell302 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 Bank Name and Address"));
            Cell cell303 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               //.Add(new Paragraph(TxtBankNameAPG2.Text + ", " + TxtBankAddressAPG2.Text).SetFont(KalimatiFont));
               .Add(new Paragraph(TxtBankNameAPG2.Text + ", " + TxtBankAddressAPG2.Text));

            //Row04------------------------------------------------------
            Cell cell04 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("E. Performance Bond Bank Guarantee"));

            //Row31------------------------------------------------------
            Cell cell311 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("31"));

            Cell cell312 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Document Reference no."));
            Cell cell313 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBRefNo.Text));

            //Row32------------------------------------------------------
            Cell cell321 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("32"));

            Cell cell322 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Deadline"));
            Cell cell323 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBDL.Text));


            //Row33------------------------------------------------------
            Cell cell331 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("33"));

            Cell cell332 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Amount"));
            Cell cell333 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBAmount.Text));

            //Row34------------------------------------------------------
            Cell cell341 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("34"));

            Cell cell342 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Minimum DeadLine"));
            Cell cell343 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBMinDL.Text));

            //Row35------------------------------------------------------
            Cell cell351 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("35"));

            Cell cell352 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Remark"));
            Cell cell353 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBRemark.Text));

            //Row36------------------------------------------------------
            Cell cell361 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("36"));

            Cell cell362 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB Bank Name and Address"));
            Cell cell363 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtBankNamePB.Text + ", " + TxtBankAddressPB.Text));
            //.Add(new Paragraph(TxtBankNamePB.Text + ", " + TxtBankAddressPB.Text).SetFont(KalimatiFont));

            //Row05------------------------------------------------------
            Cell cell05 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("F. Insurance"));

            //Row37------------------------------------------------------
            Cell cell371 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("37"));

            Cell cell372 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins Document Reference no."));
            Cell cell373 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsRefNo.Text));

            //Row38------------------------------------------------------
            Cell cell381 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("38"));

            Cell cell382 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins Deadline"));
            Cell cell383 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsDL.Text));


            //Row39------------------------------------------------------
            Cell cell391 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("39"));

            Cell cell392 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins Amount"));
            Cell cell393 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsAmount.Text));

            //Row40------------------------------------------------------
            Cell cell401 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("40"));

            Cell cell402 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins Minimum DeadLine"));
            Cell cell403 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsMinDL.Text));

            //Row41------------------------------------------------------
            Cell cell411 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("41"));

            Cell cell412 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins Remark"));
            Cell cell413 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsRemark.Text));

            //Row42------------------------------------------------------
            Cell cell421 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("42"));

            Cell cell422 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Insurance company Name and Address"));
            Cell cell423 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               //.Add(new Paragraph(TxtBankNameIns.Text + ", " + TxtBankAddressIns.Text).SetFont(KalimatiFont));
               .Add(new Paragraph(TxtBankNameIns.Text + ", " + TxtBankAddressIns.Text));


            //Row06------------------------------------------------------
            Cell cell06 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("G. Contractor's Information"));

            //Row43------------------------------------------------------
            Cell cell431 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("43"));

            Cell cell432 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contractor's Name and Address"));
            Cell cell433 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractorName.Text + ", " + TxtAddressOfContractor.Text));

            //Row44------------------------------------------------------
            /*Cell cell441 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("44"));

            Cell cell442 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contractor's Name and Address (Devanagiri)"));
            Cell cell443 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractorNameDev.Text + ", " + TxtContractorAddressDev.Text).SetFont(KalimatiFont));
            */
            //Row45------------------------------------------------------
            Cell cell451 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("44"));

            Cell cell452 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contractor's Email"));
            Cell cell453 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtEmail1.Text));

            //Row46------------------------------------------------------
            Cell cell461 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("45"));

            Cell cell462 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Contractor's Other Information"));
            Cell cell463 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtContractorOther.Text));


            //Row07------------------------------------------------------
            Cell cell07 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("H. Project Informtion"));

            //Row46------------------------------------------------------
            Cell cell471 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("46"));

            Cell cell472 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Project Description"));
            Cell cell473 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtProjectDescription.Text));

            //Row47------------------------------------------------------
            Cell cell4711 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("47"));

            Cell cell4721 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Procurement Cat and Method"));
            Cell cell4731 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtProcurementcategory.Text + " - " + TxtProcurementMethod.Text));

            //Row48------------------------------------------------------
            Cell cell481 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("48"));

            Cell cell482 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Length, Breadth and Height (m)"));
            Cell cell483 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("L = " + TxtLength.Text + ", B = " + TxtBreadth.Text + " and H = " + TxtHeight.Text));

            //Row07------------------------------------------------------
            Cell cell08 = new Cell(1, 3)
               .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.GRAY)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("I. Days Remaining from Today"));

            //Row20.1------------------------------------------------------
            Cell cellD201 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("49"));

            Cell cellD202 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG1 days remaining from Today"));
            Cell cellD203 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG1DaysRem.Text));

            //Row26.1------------------------------------------------------
            Cell cellD261 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("50"));

            Cell cellD262 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("APG2 days remaining from Today"));
            Cell cellD263 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtAPG2DaysRem.Text));


            //Row32.1------------------------------------------------------
            Cell cellD321 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("51"));

            Cell cellD322 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("PB days remaining from Today"));
            Cell cellD323 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtPBDaysRem.Text));

            //Row38.1------------------------------------------------------
            Cell cellD381 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("52"));

            Cell cellD382 = new Cell(1, 1)
               //.SetBackgroundColor(Color.Green)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph("Ins days remaining from Today"));
            Cell cellD383 = new Cell(1, 1)
               //.SetBackgroundColor(Color.GRAY)
               .SetTextAlignment(TextAlignment.LEFT)
               .Add(new Paragraph(TxtInsDaysRem.Text));

            table.AddCell(cell00); //A

            table.AddCell(cell11);
            table.AddCell(cell12);
            table.AddCell(cell13);

            table.AddCell(cell21);
            table.AddCell(cell22);
            table.AddCell(cell23);

            table.AddCell(cell31);
            table.AddCell(cell32);
            table.AddCell(cell33);

            table.AddCell(cell41);
            table.AddCell(cell42);
            table.AddCell(cell43);

            table.AddCell(cell51);
            table.AddCell(cell52);
            table.AddCell(cell53);

            table.AddCell(cell61);
            table.AddCell(cell62);
            table.AddCell(cell63);

            table.AddCell(cell71);
            table.AddCell(cell72);
            table.AddCell(cell73);

            table.AddCell(cell81);
            table.AddCell(cell82);
            table.AddCell(cell83);

            table.AddCell(cell01); //B

            table.AddCell(cell91);
            table.AddCell(cell92);
            table.AddCell(cell93);

            table.AddCell(cell101);
            table.AddCell(cell102);
            table.AddCell(cell103);

            table.AddCell(cell111);
            table.AddCell(cell112);
            table.AddCell(cell113);

            table.AddCell(cell121);
            table.AddCell(cell122);
            table.AddCell(cell123);

            table.AddCell(cell131);
            table.AddCell(cell132);
            table.AddCell(cell133);

            table.AddCell(cell141);
            table.AddCell(cell142);
            table.AddCell(cell143);

            table.AddCell(cell151);
            table.AddCell(cell152);
            table.AddCell(cell153);

            table.AddCell(cell161);
            table.AddCell(cell162);
            table.AddCell(cell163);

            table.AddCell(cell171);
            table.AddCell(cell172);
            table.AddCell(cell173);

            table.AddCell(cell181);
            table.AddCell(cell182);
            table.AddCell(cell183);

            table.AddCell(cell02); //c

            //APG1
            table.AddCell(cell191);
            table.AddCell(cell192);
            table.AddCell(cell193);

            table.AddCell(cell201);
            table.AddCell(cell202);
            table.AddCell(cell203);

            table.AddCell(cell211);
            table.AddCell(cell212);
            table.AddCell(cell213);

            table.AddCell(cell221);
            table.AddCell(cell222);
            table.AddCell(cell223);

            table.AddCell(cell231);
            table.AddCell(cell232);
            table.AddCell(cell233);

            table.AddCell(cell241);
            table.AddCell(cell242);
            table.AddCell(cell243);

            table.AddCell(cell03); //D

            //APG2
            table.AddCell(cell251);
            table.AddCell(cell252);
            table.AddCell(cell253);

            table.AddCell(cell261);
            table.AddCell(cell262);
            table.AddCell(cell263);

            table.AddCell(cell271);
            table.AddCell(cell272);
            table.AddCell(cell273);

            table.AddCell(cell281);
            table.AddCell(cell282);
            table.AddCell(cell283);

            table.AddCell(cell291);
            table.AddCell(cell292);
            table.AddCell(cell293);

            table.AddCell(cell301);
            table.AddCell(cell302);
            table.AddCell(cell303);

            table.AddCell(cell04); //E

            //PB
            table.AddCell(cell311);
            table.AddCell(cell312);
            table.AddCell(cell313);

            table.AddCell(cell321);
            table.AddCell(cell322);
            table.AddCell(cell323);

            table.AddCell(cell331);
            table.AddCell(cell332);
            table.AddCell(cell333);

            table.AddCell(cell341);
            table.AddCell(cell342);
            table.AddCell(cell343);

            table.AddCell(cell351);
            table.AddCell(cell352);
            table.AddCell(cell353);

            table.AddCell(cell361);
            table.AddCell(cell362);
            table.AddCell(cell363);

            table.AddCell(cell05); //F

            //Ins
            table.AddCell(cell371);
            table.AddCell(cell372);
            table.AddCell(cell373);

            table.AddCell(cell381);
            table.AddCell(cell382);
            table.AddCell(cell383);

            table.AddCell(cell391);
            table.AddCell(cell392);
            table.AddCell(cell393);

            table.AddCell(cell401);
            table.AddCell(cell402);
            table.AddCell(cell403);

            table.AddCell(cell411);
            table.AddCell(cell412);
            table.AddCell(cell413);

            table.AddCell(cell421);
            table.AddCell(cell422);
            table.AddCell(cell423);

            table.AddCell(cell06); //G

            table.AddCell(cell431);
            table.AddCell(cell432);
            table.AddCell(cell433);

            //table.AddCell(cell441);
            //table.AddCell(cell442);
            //table.AddCell(cell443);

            table.AddCell(cell451);
            table.AddCell(cell452);
            table.AddCell(cell453);

            table.AddCell(cell461);
            table.AddCell(cell462);
            table.AddCell(cell463);

            table.AddCell(cell07); //H

            table.AddCell(cell471);
            table.AddCell(cell472);
            table.AddCell(cell473);

            table.AddCell(cell4711);
            table.AddCell(cell4721);
            table.AddCell(cell4731);

            table.AddCell(cell481);
            table.AddCell(cell482);
            table.AddCell(cell483);

            table.AddCell(cell08); //I


            table.AddCell(cellD201);
            table.AddCell(cellD202);
            table.AddCell(cellD203);

            table.AddCell(cellD261);
            table.AddCell(cellD262);
            table.AddCell(cellD263);

            table.AddCell(cellD321);
            table.AddCell(cellD322);
            table.AddCell(cellD323);

            table.AddCell(cellD381);
            table.AddCell(cellD382);
            table.AddCell(cellD383);

            document.Add(table);

            document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

            Paragraph header2 = new Paragraph();
            header2.Add("Bill of Contract ID: " + TxtContractID.Text + " at " + TxtCurrentStatus.Text)
                .SetTextAlignment(TextAlignment.CENTER)
                //.SetFontColor(iText.Kernel.Colors.ColorConstants.RED)
                .SetFontSize(14);
            //.SetFont(KalimatiFont);
            document.Add(header2);

            float[] colwidth = new float[] { 15f, 15f, 15f, 15f, 15f, 15f, 15f, 15f };

            iText.Layout.Element.Table table1 = PDFTableFromDGV(dataGridView1, colwidth);
            document.Add(table1);

            document.Close();
            MessageBox.Show("Pdf Created Successfully.", "Create Pdf");
        }

        private void BtnCreateAllPdf_Click(object sender, EventArgs e)
        {

        }
        private iText.Layout.Element.Table PDFTableFromDGV(DataGridView dgv, float[] cloumnwidth)
        {
            // Getting Rows & Columns Counts
            int dgvrowcount = dgv.Rows.Count - 1;//12
            int dgvcolumncount = dgv.Columns.Count;//8
            //MessageBox.Show(dgvrowcount + " and " + dgvcolumncount, "row and ocl");
            string[,] datagridcontent = new string[15, 10];

            // Set The Table like new float [] {15f, 15f, 15f, 15f, 15f }
            iText.Layout.Element.Table table = new iText.Layout.Element.Table(cloumnwidth);
            table.SetWidth(iText.Layout.Properties.UnitValue.CreatePercentValue(100));

            // Print The DGV Header To Table Header
            for (int i = 0; i < dgvcolumncount; i++)
            {
                Cell headerCells = new Cell()
                              .SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
                              .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
                //headerCells.SetNextRenderer(new RoundedCornersCellRenderer(headerCells));
                var gteCell = headerCells.Add(new Paragraph(dgv.Columns[i].HeaderText));

                table.AddHeaderCell(gteCell);
            }

            // Print The DGV Cells To Table Cells
            for (int i = 0; i < 12; i++) //dgvrowcount
            {
                for (int c = 0; c < 8; c++) //dgvcolumncount
                {
                    datagridcontent[i, c] = dataGridView1.Rows[i].Cells[c].Value.ToString();

                    Cell gteCell = new Cell(1, 1)
                       //.SetBackgroundColor(Color.Green)
                       .SetTextAlignment(TextAlignment.LEFT)
                       .Add(new Paragraph(datagridcontent[i, c]));
                    table.AddCell(gteCell);
                }
            }

            return table;
        }

        private void ComboBoxPE_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtPE.Text = ComboBoxPE.Text;
        }
        private void CopyAlltoClipboard(DataGridView DGV)
        {
            DGV.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DGV.MultiSelect = true;
            DGV.SelectAll();
            DataObject dataObj = DGV.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void BtnExportAllToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                CopyAlltoClipboard(dataGridView2);
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                ((Excel.Range)xlWorkSheet.Cells[1, 1]).Value = "Record " + DateTime.Now.ToString("yyyy/MM/dd_HH:mm:ss");

                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[5, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                // xlWorkBook.Close();
                //  xlexcel.Quit();
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlWorkSheet);

                MessageBox.Show("Export Completed Sucessfully.");

            }
            catch
            {

            }
        }

        private void exportToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                CopyAlltoClipboard(dataGridView1);
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                ((Excel.Range)xlWorkSheet.Cells[1, 1]).Value = "Record " + DateTime.Now.ToString("yyyy/MM/dd_HH:mm:ss");


                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[5, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                // xlWorkBook.Close();
                //  xlexcel.Quit();
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlWorkSheet);

                MessageBox.Show("Export Completed Sucessfully.");

            }
            catch
            {


            }
        }

        private void BtnCreateDocx_Click(object sender, EventArgs e)
        {


        }

        public static DataTable DataGridView_To_Datatable(DataGridView dg)
        {
            DataTable ExportDataTable = new DataTable();
            foreach (DataGridViewColumn col in dg.Columns)
            {
                ExportDataTable.Columns.Add(col.Name);
            }
            foreach (DataGridViewRow row in dg.Rows)
            {
                DataRow dRow = ExportDataTable.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                ExportDataTable.Rows.Add(dRow);
            }
            return ExportDataTable;
        }

        private void saveToExcelxlsxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (TxtFY.Text == "" || TxtContractID.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "")
                {
                    TxtLog.Text += "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                    TxtLog.Text = Environment.NewLine;
                    TxtBillLog.Text = "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                    TxtBillLog.Text = Environment.NewLine;
                }
                else
                {
                    //string BillFileName;

                    //string ThisContractID, ThisWard;
                    CreateAccessProjectFolders();

                    //if (dataGridView1.Rows.Count > 1)
                    //{
                    DataTable table = DataGridView_To_Datatable(dataGridView1);
                    var name1 = table.Rows[0][1];
                    //MessageBox.Show("Data is Converted!");
                    //}


                    //DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
                    var memoryStream = new MemoryStream();
                    //string filename = "Result.xlsx";
                    string filename = EventHistoryFolder + "\\Bill.xlsx";
                    using (var fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
                    {
                        IWorkbook workbook = new XSSFWorkbook();
                        ISheet excelSheet = workbook.CreateSheet("Sheet1");

                        List<String> columns = new List<string>();
                        IRow row = excelSheet.CreateRow(0);
                        int columnIndex = 0;

                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);
                            row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                            columnIndex++;
                        }

                        int rowIndex = 1;
                        foreach (DataRow dsrow in table.Rows)
                        {
                            row = excelSheet.CreateRow(rowIndex);
                            int cellIndex = 0;
                            foreach (String col in columns)
                            {
                                row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                                cellIndex++;
                            }

                            rowIndex++;
                        }
                        workbook.Write(fs);
                    }
                    MessageBox.Show("Contract and Bill saved to Excel", "Save to Excel");
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void BtnSave2Excel_Click(object sender, EventArgs e)
        {
            try
            {
                //if (dataGridView1.Rows.Count > 1)
                //{
                DataTable table = DataGridView_To_Datatable(dataGridView2);
                var name1 = table.Rows[0][1];
                //MessageBox.Show("Data is Converted!");
                //}


                //DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
                var memoryStream = new MemoryStream();
                string filename = "AllRecord.xlsx";
                //string filename = EventHistoryFolder + "\\Bill.xlsx";
                using (var fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Sheet1");

                    List<String> columns = new List<string>();
                    IRow row = excelSheet.CreateRow(0);
                    int columnIndex = 0;

                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);
                        row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                        columnIndex++;
                    }

                    int rowIndex = 1;
                    foreach (DataRow dsrow in table.Rows)
                    {
                        row = excelSheet.CreateRow(rowIndex);
                        int cellIndex = 0;
                        foreach (String col in columns)
                        {
                            row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                            cellIndex++;
                        }

                        rowIndex++;
                    }
                    workbook.Write(fs);
                }
                MessageBox.Show("All Records saved to Excel", "Save to Excel");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void ComboBoxCurrentStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtCurrentStatus.Text = ComboBoxCurrentStatus.Text;
        }

        public void SetColorofInputCells()
        {
            dataGridView1.Rows[0].Cells[2].Style.BackColor = Color.Yellow; //PS Estimate
            dataGridView1.Rows[1].Cells[2].Style.BackColor = Color.Yellow; //Subtotal Estimate

            dataGridView1.Rows[1].Cells[3].Style.BackColor = Color.Yellow; //Subtotal Contract
            dataGridView1.Rows[9].Cells[3].Style.BackColor = Color.Yellow; //AP1 Contract
            dataGridView1.Rows[10].Cells[3].Style.BackColor = Color.Yellow; //AP2 Contract

            dataGridView1.Rows[0].Cells[6].Style.BackColor = Color.Yellow; //PS Up2ThisBill
            dataGridView1.Rows[1].Cells[6].Style.BackColor = Color.Yellow; //Subtotal Up2ThisBill
            dataGridView1.Rows[9].Cells[6].Style.BackColor = Color.Yellow; //Deduct AP Up2ThisBill

        }



        private void BtnThis2Previous_Click(object sender, EventArgs e)
        {
            try
            {
                int[] InRowIndex = new int[] //Total inrows = 8, except row 3,7,8,11
                {
                    0, //PS
                    1, //Subtotal
                    3, //VAT Amount
                    7, //Tota (A+B+D)
                    8, //Toal incl. Contingencies
                    9,  //AP Deduction %
                    12, //All Deduction amount
                    13  //Net payable
                };

                //Input data from grid
                for (int j = 0; j < (rowsAmtGridBill); j++) //rowsAmtGrid = 114 i.e. j = 0 to 13 for bill
                {
                    dataGridView1.Rows[j].Cells[5].Value = dataGridView1.Rows[j].Cells[6].Value;
                }

                TxtBillLog.Text = "Recent: Data of Up2This bill transfered to Up2Previous Successfully !";

                /*//Empty Input datagrid
                for (int j = 0; j < (rowsAmtGrid - n_of_calc_Row-1); j++) //rowsAmtGrid = 12 n_of_calc_Row = 4 i.e. j = 0 to 6
                {
                    dataGridView1.Rows[InRowIndex[j]].Cells[6].Value = "";
                }*/
            }
            catch
            {

            }
        }

        private void BtnSave2Txt_Click(object sender, EventArgs e)
        {
            try
            {
                if (TxtFY.Text == "" || TxtContractID.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "")
                {
                    TxtLog.Text += "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                    TxtLog.Text = Environment.NewLine;
                    TxtBillLog.Text = "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                    TxtBillLog.Text = Environment.NewLine;
                }
                else
                {
                    DialogResult dr = MessageBox.Show("Are you sure, you want to Save Contract Bill?", "Save Bill to Text File", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                    {
                        //String[,] sd = new String[rowsAmtGrid, colsAmtGrid]; //[12,8]

                        string SaveBillintextFile, BillFileName, SaveLastBillInTxt;

                        //string ThisContractID, ThisWard;
                        CreateAccessProjectFolders();


                        SaveBillintextFile = "";
                        SaveLastBillInTxt = "";
                        SaveBillintextFile += Environment.NewLine;
                        SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                        SaveBillintextFile += Environment.NewLine;

                        SaveBillintextFile += "--------------------------------------";
                        SaveBillintextFile += Environment.NewLine;

                        /*for (int col = 0; col < dataGridView1.ColumnCount; col++)
                        {
                            SaveBillintextFile += Convert.ToString(dataGridView1.Columns[col].HeaderText);
                        }*/

                        SaveBillintextFile += "--------------------------------------";
                        SaveBillintextFile += Environment.NewLine;

                        for (int i = 0; i < rowsAmtGridBill; i++)
                        {
                            for (int j = 0; j < colsAmtGrid; j++)
                            {
                                SaveBillintextFile += dataGridView1.Rows[i].Cells[j].Value;
                                SaveBillintextFile += "\t";

                                SaveLastBillInTxt += dataGridView1.Rows[i].Cells[j].Value;
                                SaveLastBillInTxt += "\t";

                            }
                            SaveBillintextFile += Environment.NewLine;
                            SaveLastBillInTxt += Environment.NewLine;
                        }

                        BillFileName = EventHistoryFolder + "\\Bill.txt";
                        //using (StreamWriter sw = File.AppendText(@".\EventHistory.\Bill\" + BillFileName))
                        using (StreamWriter sw = File.AppendText(BillFileName))
                        {
                            sw.WriteLine(SaveBillintextFile);
                        }
                        TxtBillLog.Text = "Recent: Appended to Text Successful at " + BillFileName;
                        TxtBillLog.Text = Environment.NewLine;

                        BillFileName = LastEventFolder + "\\LastBill.txt";
                        //using (StreamWriter swl = new StreamWriter(@".\LastEvent\LastBill.txt"))
                        using (StreamWriter swl = new StreamWriter(BillFileName))
                        {
                            swl.Write(SaveLastBillInTxt);
                        }

                        TxtBillLog.Text += "Recent: Last bill Saved to Text Successfully at " + BillFileName;
                        TxtBillLog.Text = Environment.NewLine;
                    }
                    else if (dr == DialogResult.No)
                    {
                        //do nothing
                        TxtBillLog.Text = "Recent: Save to Text cancelled !";
                    }
                }

                SaveLetterTippaniInfo();

            }
            catch
            {

            }
        }

        private void SaveLetterTippaniInfo()
        {
            try
            {
                if (TxtFY.Text == "" || TxtContractID.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "")
                {
                    TxtLog.Text += "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to save letter info.";
                    TxtLog.Text = Environment.NewLine;
                }
                else
                {
                    string SaveAPGTippaniInfo, FileName, SaveLetterInfo;

                    CreateAccessProjectFolders();


                    SaveAPGTippaniInfo = "";
                    SaveLetterInfo = "";

                    int rows, cols;

                    //Letter info AD and BS
                    rows = 7;
                    cols = 3;
                    SaveLetterInfo += "Description\tDate(AD)\tDate(BS)";
                    SaveLetterInfo += Environment.NewLine;
                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            SaveLetterInfo += dataGridView3.Rows[i].Cells[j].Value;
                            SaveLetterInfo += "\t";

                        }
                        SaveLetterInfo += Environment.NewLine;
                    }

                    FileName = LastEventFolder + "\\LetterDatesAD_BS.txt";
                    //using (StreamWriter swl = new StreamWriter(@".\LastEvent\LastBill.txt"))
                    using (StreamWriter swl = new StreamWriter(FileName))
                    {
                        swl.Write(SaveLetterInfo);
                    }

                    TxtLog.Text += "Recent: Letter tippani dates AD-BS info saved at " + FileName;
                    TxtLog.Text = Environment.NewLine;


                    //APG tippani Info
                    rows = 3;
                    cols = 3;
                    SaveAPGTippaniInfo += "Description\tAPG1\tAPG2";
                    SaveAPGTippaniInfo += Environment.NewLine;
                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            SaveAPGTippaniInfo += dataGridView4.Rows[i].Cells[j].Value;
                            SaveAPGTippaniInfo += "\t";

                        }
                        SaveAPGTippaniInfo += Environment.NewLine;
                    }

                    FileName = LastEventFolder + "\\AP_Tippani_Info.txt";
                    //using (StreamWriter swl = new StreamWriter(@".\LastEvent\LastBill.txt"))
                    using (StreamWriter swl = new StreamWriter(FileName))
                    {
                        swl.Write(SaveAPGTippaniInfo);
                    }

                    TxtLog.Text += "Recent: AP Tippani info saved at " + FileName;
                    TxtLog.Text = Environment.NewLine;

                }

            }
            catch
            {

            }
        }

        private void ReadLetterTippaniInfo()
        {
            try
            {
                string FileName;
                CreateAccessProjectFolders();
                FileName = LastEventFolder + "\\LetterDatesAD_BS.txt";
                LoadTxtToDatagridview(dataGridView3, FileName, 1, 3);//index of first line of .txt file is 1
                TxtBillLog.Text = "Recent: Read letter date from Text Successfully: " + FileName;

                FileName = LastEventFolder + "\\AP_Tippani_Info.txt";
                LoadTxtToDatagridview(dataGridView4, FileName, 1, 3);

                TxtBillLog.Text = "Recent: Read AP tippani info from Text Successfully: " + FileName;
            }
            catch
            {

            }
        }

        private void BtnReadfromTxt_Click(object sender, EventArgs e)
        {
            try
            {
                string[] ReadingText = new string[15];
                string BillFilenName;
                CreateAccessProjectFolders();
                string line;
                line = "";
                BillFilenName = LastEventFolder + "\\LastBill.txt";
                //Pass the file path and file name to the StreamReader constructor
                //StreamReader sr = new StreamReader(@".\LastEvent\LastBill.txt");
                StreamReader sr = new StreamReader(BillFilenName);
                //Read the first line of text
                line = sr.ReadLine();
                ReadingText[0] = line;
                //Continue to read until you reach end of file
                int i = 1;
                while (line != null)
                {
                    //Read the next line
                    line = sr.ReadLine();
                    ReadingText[i] = line;
                    i++;
                }
                //close the file
                sr.Close();

                //load data to datagridview by splitting by tab character
                for (int row = 0; row < 12; row++)
                {
                    string[] splittedtext = ReadingText[row].Split("\t");
                    for (int col = 0; col < 8; col++)
                    {
                        dataGridView1.Rows[row].Cells[col].Value = splittedtext[col];
                    }
                }
                TxtBillLog.Text = "Recent: Read from Text Successfully: " + BillFilenName;

                ReadLetterTippaniInfo();
            }
            catch
            {

            }
        }

        private void Fun_CreateProjectFolder()
        {
            try
            {
                if (TxtFY.Text == "" || TxtContractID.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "")
                {
                    TxtLog.Text += "Either Fiscal Year or Contract ID or Ward or Project Type is Empty. Please fill to continue.";
                    TxtLog.Text += Environment.NewLine;
                }
                else
                {
                    CreateAccessProjectFolders();

                    if (!Directory.Exists(Project_Folders))
                    {
                        Directory.CreateDirectory(Project_Folders);
                    }

                    //create individual contract folder 
                    if (!Directory.Exists(ThisContractFolder))
                    {
                        Directory.CreateDirectory(ThisContractFolder);
                    }

                    //create EventHistory folder
                    if (!Directory.Exists(EventHistoryFolder))
                    {
                        Directory.CreateDirectory(EventHistoryFolder);
                    }

                    //create LastEvent folder
                    if (!Directory.Exists(LastEventFolder))
                    {
                        Directory.CreateDirectory(LastEventFolder);
                    }

                    //---------------------------------------------------------
                    //write general infomation
                    string SaveBillintextFile, BillFileName;

                    SaveBillintextFile = "";
                    SaveBillintextFile += Environment.NewLine;
                    SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "General Information of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "1. Project ID\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtProjectID.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "2. Fiscal Year\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtFY.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "3. Contract ID\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtContractID.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "4. Contract Name\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtContractName.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "5. Contract Budget\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtContractBudget.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "6. Location and Ward\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtLocation.Text + "-" + TxtWard.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "7. Budget Type\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtBudgetType.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "9. Project Type\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtProjectType.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "Important Event Dates of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "10. Current Status\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtCurrentStatus.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "11. Notice Issued\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtNoticeIssued.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "12. LOI\t\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtLOI.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "13. LOA\t\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtLOA.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "14. Contract Agreement\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtContractAgreement.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "15. Work Permit\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtWorkPermit.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "16. Work Complete\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtWorkComplete.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "17. Running Bill\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtRunningBill.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "18. Final Bill\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtFinalBill.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "19. Contract Day remaining\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtDaysRem.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "Procurement Info";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "20. Procurement Category\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtProcurementcategory.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "21. Procurement Method\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtProcurementMethod.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "22. Public Entity\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPE.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "Amount Summary Info";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "23. Total Estimated Amount\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtTotalEstimatedAmount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "24. Total Contract Amount\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtTotalContractAmount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "25. Total Final Bill Amount\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtTotalFinalBillAmount.Text;
                    SaveBillintextFile += Environment.NewLine;


                    BillFileName = EventHistoryFolder + "\\General_Info.txt";
                    using (StreamWriter swl = new StreamWriter(BillFileName))
                    {
                        swl.Write(SaveBillintextFile);
                    }


                    //Append APG1
                    SaveBillintextFile = "";
                    SaveBillintextFile += Environment.NewLine;
                    SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "APG1 of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "1. APG1 Doc Ref. No.:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG1RefNo.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "2. APG1 Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG1DL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "3. APG1 Amount:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG1Amount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "4. APG1 Min Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG1MinDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "5. Remark:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG1Remark.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "6. Bank Name and Address:\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtBankNameAPG1.Text + "," + TxtBankAddressAPG1.Text;
                    SaveBillintextFile += Environment.NewLine;

                    BillFileName = EventHistoryFolder + "\\APG1.txt";
                    using (StreamWriter sw = File.AppendText(BillFileName))
                    {
                        sw.WriteLine(SaveBillintextFile);
                    }

                    //Append APG2
                    SaveBillintextFile = "";
                    SaveBillintextFile += Environment.NewLine;
                    SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "APG2 of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "1. APG2 Doc Ref. No.:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG2RefNo.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "2. APG2 Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG2DL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "3. APG2 Amount:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG2Amount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "4. APG2 Min Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG2MinDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "5. Remark:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtAPG2Remark.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "6. Bank Name and Address:\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtBankNameAPG2.Text + "," + TxtBankAddressAPG2.Text;
                    SaveBillintextFile += Environment.NewLine;

                    BillFileName = EventHistoryFolder + "\\APG2.txt";
                    using (StreamWriter sw = File.AppendText(BillFileName))
                    {
                        sw.WriteLine(SaveBillintextFile);
                    }

                    //Append PB
                    SaveBillintextFile = "";
                    SaveBillintextFile += Environment.NewLine;
                    SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "PB of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "1. PB Doc Ref. No.:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPBRefNo.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "2. PB Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPBDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "3. PB Amount:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPBAmount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "4. PB Min Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPBMinDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "5. Remark:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtPBRemark.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "6. Bank Name and Address:\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtBankNamePB.Text + "," + TxtBankAddressPB.Text;
                    SaveBillintextFile += Environment.NewLine;

                    BillFileName = EventHistoryFolder + "\\PB.txt";
                    using (StreamWriter sw = File.AppendText(BillFileName))
                    {
                        sw.WriteLine(SaveBillintextFile);
                    }

                    //Append Ins
                    SaveBillintextFile = "";
                    SaveBillintextFile += Environment.NewLine;
                    SaveBillintextFile += "ModifiedDate:" + DateTime.Now.ToString("F");
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "Ins of the Contract";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "-----------------------------------------------------------";
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "1. Ins Doc Ref. No.:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtInsRefNo.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "2. Ins Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtInsDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "3. Ins Amount:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtInsAmount.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "4. Ins Min Deadline:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtInsMinDL.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "5. Remark:\t\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtInsRemark.Text;
                    SaveBillintextFile += Environment.NewLine;

                    SaveBillintextFile += "6. Bank Name and Address:\t:";
                    SaveBillintextFile += "\t";
                    SaveBillintextFile += TxtBankNameIns.Text + "," + TxtBankAddressIns.Text;
                    SaveBillintextFile += Environment.NewLine;

                    BillFileName = EventHistoryFolder + "\\Insurance.txt";
                    using (StreamWriter sw = File.AppendText(BillFileName))
                    {
                        sw.WriteLine(SaveBillintextFile);
                    }
                }

            }
            catch
            {

            }
        }

        private void CreateAccessProjectFolders()
        {
            Cur_Dir = Environment.CurrentDirectory;
            FYFolder = TxtFY.Text;

            Ward = TxtWard.Text;

            if (Ward == "")
            {
                Ward = "0";
            }
            Contract_ID = TxtContractID.Text;
            if (Contract_ID == "")
            {
                Contract_ID = "Other_Contract";
            }
            Contract_ID = Contract_ID.Replace("/", "-");
            Contract_ID = Contract_ID.Replace("\\", "-");
            Project_Type = TxtProjectType.Text;
            if (Project_Type == "")
            {
                Project_Type = "Other_Project";
            }
            Project_Folders = Cur_Dir + "\\ProjectFolders\\" + FYFolder + "\\" + Project_Type;
            ThisContractFolder = Project_Folders + "\\" + Ward + " " + Contract_ID;
            EventHistoryFolder = ThisContractFolder + "\\EventHistory";
            LastEventFolder = ThisContractFolder + "\\LastEvent";
        }

        private void BtnResetBill_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            GenerateAmountDataGridFromText();
            //GenerateAmountDataGrid();
            SetColorofInputCells();

            dataGridView1.Rows[0].Cells[6].Style.ForeColor = Color.Black;
            dataGridView1.Rows[1].Cells[6].Style.ForeColor = Color.Black;

            dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.White;

            LblAmountValidity.Text = "Click calculate to check Amount Validity";
            LblAmountValidity.ForeColor = Color.Black;

            TxtBillLog.Text = "Recent: Reset Successfully !";
        }

        private void BtnCalcBill_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows[0].Cells[3].Value = dataGridView1.Rows[0].Cells[2].Value.ToString();

                double PS1, ST1, VAT_Per, Contingencies, VAT_Amount, GT_exclCont, GT_inclCont;
                double AP1, AP2, AP_Total, Ap_ded_amount;
                double AP_deduction_Per, Deduction_Per, Net_Pay;

                int[] CalcColIndex = new int[]
               {
                    2, //Estimate Column
                    3, //Contract Column
                    6  //Amount up2 this bill column
               };

                for (int k = 0; k < 3; k++)
                {
                    int idx = CalcColIndex[k];

                    PS1 = Convert.ToDouble(dataGridView1.Rows[0].Cells[idx].Value);
                    ST1 = Convert.ToDouble(dataGridView1.Rows[1].Cells[idx].Value);
                    VAT_Per = Convert.ToDouble(dataGridView1.Rows[2].Cells[idx].Value);
                    VAT_Amount = (Math.Round(ST1 * VAT_Per / 100.0, 2));

                    dataGridView1.Rows[3].Cells[idx].Value = VAT_Amount.ToString();
                    GT_exclCont = Math.Round(PS1 + ST1 + VAT_Amount, 2);
                    dataGridView1.Rows[7].Cells[idx].Value = GT_exclCont.ToString();

                    Contingencies = 0;
                    for (int i = 4; i <= 5; i++)
                    {
                        Contingencies += Convert.ToDouble(dataGridView1.Rows[i].Cells[idx].Value);
                    }
                    GT_inclCont = Math.Round(Contingencies / 100.0 * ST1 + GT_exclCont, 2);
                    dataGridView1.Rows[8].Cells[idx].Value = GT_inclCont.ToString();

                    if (idx == 3)
                    {
                        AP1 = Convert.ToDouble(dataGridView1.Rows[9].Cells[3].Value);
                        AP2 = Convert.ToDouble(dataGridView1.Rows[10].Cells[3].Value);
                        AP_Total = AP1 + AP2;
                        dataGridView1.Rows[11].Cells[3].Value = AP_Total.ToString();
                    }

                    if (idx == 6)
                    {
                        AP_deduction_Per = Convert.ToDouble(dataGridView1.Rows[9].Cells[6].Value);
                        Deduction_Per = Convert.ToDouble(dataGridView1.Rows[10].Cells[6].Value);

                        AP_Total = Convert.ToDouble(dataGridView1.Rows[11].Cells[3].Value);
                        Ap_ded_amount = AP_Total * AP_deduction_Per / 100.0;
                        Net_Pay = Math.Round(GT_exclCont - (Deduction_Per / 100.0 * ST1) - Ap_ded_amount, 2);
                        dataGridView1.Rows[11].Cells[6].Value = Net_Pay.ToString();
                    }

                }

                //calculate for ThisBillOnly
                for (int j = 0; j < rowsAmtGrid; j++)
                {
                    if (j == 2 || j == 4 || j == 5 || j == 6 || j == 10)
                    {
                        dataGridView1.Rows[j].Cells[7].Value = dataGridView1.Rows[j].Cells[6].Value;
                    }
                    else
                    {
                        double num1, num2;
                        num1 = Convert.ToDouble(dataGridView1.Rows[j].Cells[6].Value);
                        num2 = Convert.ToDouble(dataGridView1.Rows[j].Cells[5].Value);
                        dataGridView1.Rows[j].Cells[7].Value = Math.Round(num1 - num2, 2).ToString();
                    }
                }

                //checking if evaluated amount is greater than contract amount
                double PS_CA, ST_CA, PS_E, ST_E;
                PS_CA = Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value);
                ST_CA = Convert.ToDouble(dataGridView1.Rows[1].Cells[3].Value);
                //T_CA = Convert.ToSingle(dataGridView1.Rows[0].Cells[3].Value);

                PS_E = Convert.ToDouble(dataGridView1.Rows[0].Cells[6].Value);
                ST_E = Convert.ToDouble(dataGridView1.Rows[1].Cells[6].Value);
                //T_E = Convert.ToSingle(dataGridView1.Rows[0].Cells[3].Value);
                if (PS_CA < PS_E)
                {
                    dataGridView1.Rows[0].Cells[6].Style.ForeColor = Color.DarkRed;
                    //dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.Violet;
                }
                else if (PS_CA >= PS_E)
                {
                    dataGridView1.Rows[0].Cells[6].Style.ForeColor = Color.ForestGreen;
                    //dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.LightGreen;
                }
                if (ST_CA < ST_E)
                {
                    dataGridView1.Rows[1].Cells[6].Style.ForeColor = Color.DarkRed;
                    //dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.LightSalmon;


                }
                else if (ST_CA >= ST_E)
                {
                    dataGridView1.Rows[1].Cells[6].Style.ForeColor = Color.ForestGreen;
                    //dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.LightGreen;
                }

                if (PS_CA < PS_E || ST_CA < ST_E)
                {
                    LblAmountValidity.Text = "Review: PS/Sub-total of bill is greater than that of Contract!";
                    LblAmountValidity.ForeColor = Color.Red;

                    dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.LightSalmon;
                }
                else if (PS_CA >= PS_E || ST_CA >= ST_E)
                {
                    LblAmountValidity.Text = "OK: PS/Sub-total of bill is less than that of Contract!";
                    LblAmountValidity.ForeColor = Color.ForestGreen;

                    dataGridView1.Rows[7].Cells[6].Style.BackColor = Color.LightGreen;
                }

                //sending est, contract, bill amount to textboxes
                TxtTotalEstimatedAmount.Text = dataGridView1.Rows[7].Cells[2].Value.ToString();
                TxtTotalContractAmount.Text = dataGridView1.Rows[7].Cells[3].Value.ToString();
                TxtTotalFinalBillAmount.Text = dataGridView1.Rows[7].Cells[6].Value.ToString();

                TxtBillLog.Text = "Recent: Calculated Successfully !";
            }
            catch
            {

            }

        }

        private void aboutToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            FrmAbout fabout = new FrmAbout();
            fabout.Show();
        }

        private void exitToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            Close();
        }

        private void addToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            Fun_Add(sender, e);
        }

        private void displayToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            if (TxtProjectID.Text == "")
            {
                TxtLog.AppendText("Enter Project ID to Display");
                TxtLog.AppendText(Environment.NewLine);
            }
            else
            {
                SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
                ConnectDb.Open();

                string query = "SELECT * FROM ContractTable where ProjectID = '" + TxtProjectID.Text + "'";

                SQLiteDataAdapter DataAdptr = new SQLiteDataAdapter(query, ConnectDb);

                DataTable Dt = new DataTable();
                DataAdptr.Fill(Dt);
                //string value;
                foreach (DataRow row in Dt.Rows) //there is only one row here
                {
                    TxtFY.Text = row[1].ToString();
                    TxtContractID.Text = row[2].ToString();
                    TxtContractName.Text = row[3].ToString();
                    TxtContractBudget.Text = row[4].ToString();
                    TxtWard.Text = row[5].ToString();
                    TxtProjectType.Text = row[6].ToString();
                    TxtBudgetType.Text = row[7].ToString();
                    TxtLocation.Text = row[8].ToString();

                    TxtAPG1RefNo.Text = row[9].ToString();
                    TxtAPG1DL.Text = row[10].ToString();
                    TxtAPG1Amount.Text = row[11].ToString();
                    TxtAPG1MinDL.Text = row[12].ToString();
                    TxtAPG1Remark.Text = row[13].ToString();

                    TxtAPG2RefNo.Text = row[14].ToString();
                    TxtAPG2DL.Text = row[15].ToString();
                    TxtAPG2Amount.Text = row[16].ToString();
                    TxtAPG2MinDL.Text = row[17].ToString();
                    TxtAPG2Remark.Text = row[18].ToString();

                    TxtPBRefNo.Text = row[19].ToString();
                    TxtPBDL.Text = row[20].ToString();
                    TxtPBAmount.Text = row[21].ToString();
                    TxtPBMinDL.Text = row[22].ToString();
                    TxtPBRemark.Text = row[23].ToString();

                    TxtInsRefNo.Text = row[24].ToString();
                    TxtInsDL.Text = row[25].ToString();
                    TxtInsAmount.Text = row[26].ToString();
                    TxtInsMinDL.Text = row[27].ToString();
                    TxtInsRemark.Text = row[28].ToString();

                    TxtCurrentStatus.Text = row[29].ToString();
                    TxtNoticeIssued.Text = row[30].ToString();
                    TxtLOI.Text = row[31].ToString();
                    TxtLOA.Text = row[32].ToString();
                    TxtContractAgreement.Text = row[33].ToString();
                    TxtWorkPermit.Text = row[34].ToString();
                    TxtWorkComplete.Text = row[35].ToString();
                    TxtRunningBill.Text = row[36].ToString();
                    TxtFinalBill.Text = row[37].ToString();
                    TxtDaysRem.Text = row[38].ToString();

                    TxtContractorName.Text = row[39].ToString();
                    TxtAddressOfContractor.Text = row[40].ToString();
                    TxtEmail1.Text = row[41].ToString();
                    TxtContractorOther.Text = row[42].ToString();

                    TxtProjectDescription.Text = row[43].ToString();
                    TxtLength.Text = row[44].ToString();
                    TxtBreadth.Text = row[45].ToString();
                    TxtHeight.Text = row[46].ToString();

                    TxtContractorNameDev.Text = row[47].ToString();
                    TxtContractorAddressDev.Text = row[48].ToString();

                    TxtAPG1DaysRem.Text = row[49].ToString();
                    TxtAPG2DaysRem.Text = row[50].ToString();
                    TxtPBDaysRem.Text = row[51].ToString();
                    TxtInsDaysRem.Text = row[52].ToString();

                    TxtBankNameAPG1.Text = row[53].ToString();
                    TxtBankNameAPG2.Text = row[54].ToString();
                    TxtBankNamePB.Text = row[55].ToString();
                    TxtBankNameIns.Text = row[56].ToString();

                    TxtBankAddressAPG1.Text = row[57].ToString();
                    TxtBankAddressAPG2.Text = row[58].ToString();
                    TxtBankAddressPB.Text = row[59].ToString();
                    TxtBankAddressIns.Text = row[60].ToString();

                    TxtProcurementcategory.Text = row[61].ToString();
                    TxtProcurementMethod.Text = row[62].ToString();
                    TxtTotalEstimatedAmount.Text = row[63].ToString();
                    TxtTotalContractAmount.Text = row[64].ToString();
                    TxtTotalFinalBillAmount.Text = row[65].ToString();
                    TxtPE.Text = row[66].ToString();

                    TxtFL_PBRef.Text = row[67].ToString();
                    TxtFL_PBDeadline.Text = row[68].ToString();
                    TxtFL_PBAmount.Text = row[69].ToString();
                    TxtFL_PBMinDL.Text = row[70].ToString();
                    TxtFL_PBRemark.Text = row[71].ToString();
                    TxtFL_BankNamePB.Text = row[72].ToString();
                    TxtFL_BankAddressPB.Text = row[73].ToString();
                    TxtPBFLDaysRem.Text = row[74].ToString();

                }
                ConnectDb.Close();

                BtnReadfromTxt_Click(sender, e);

                //days remaining from today
                int rem_days;
                if (TxtDaysRem.Text == "")
                {
                    TxtDaysRem.Text = 0.ToString();
                }
                rem_days = Convert.ToInt32(TxtDaysRem.Text);

                if (rem_days > 0)
                {
                    TxtWorkComplete.BackColor = Color.LightGreen;
                    TxtDateAnalysis.ForeColor = Color.ForestGreen;
                    TxtDaysRem.ForeColor = Color.ForestGreen;
                    TxtDateAnalysis.Text = "OK." + rem_days + " days remaining for completion.";
                }
                else if (rem_days <= 0)
                {
                    TxtWorkComplete.BackColor = Color.LightCoral;
                    TxtDateAnalysis.ForeColor = Color.Red;
                    TxtDaysRem.ForeColor = Color.Red;
                    TxtDateAnalysis.Text = "REVIEW. " + rem_days + " days past Deadline.";
                }

                //Guarantee
                if (TxtAPG1Remark.Text == "Valid")
                {
                    TxtAPG1Remark.ForeColor = Color.ForestGreen;
                }
                else if (TxtAPG1Remark.Text == "Review")
                {
                    TxtAPG1Remark.ForeColor = Color.Red;
                }

                if (TxtAPG2Remark.Text == "Valid")
                {
                    TxtAPG2Remark.ForeColor = Color.ForestGreen;
                }
                else if (TxtAPG2Remark.Text == "Review")
                {
                    TxtAPG2Remark.ForeColor = Color.Red;
                }

                if (TxtPBRemark.Text == "Valid")
                {
                    TxtPBRemark.ForeColor = Color.ForestGreen;
                }
                else if (TxtPBRemark.Text == "Review")
                {
                    TxtPBRemark.ForeColor = Color.Red;
                }

                if (TxtInsRemark.Text == "Valid")
                {
                    TxtInsRemark.ForeColor = Color.ForestGreen;
                }
                else if (TxtInsRemark.Text == "Review")
                {
                    TxtInsRemark.ForeColor = Color.Red;
                }

                if (TxtFL_PBRemark.Text == "Valid")
                {
                    TxtFL_PBRemark.ForeColor = Color.ForestGreen;
                }
                else if (TxtFL_PBRemark.Text == "Review")
                {
                    TxtFL_PBRemark.ForeColor = Color.Red;
                }
                //checking APG,PB,Ins date from Today
                //APG1
                float tempdays;
                if (TxtAPG1DaysRem.Text == "")
                {
                    TxtAPG1DaysRem.Text = 0.ToString();
                }
                tempdays = Convert.ToSingle(TxtAPG1DaysRem.Text);
                if (tempdays > 7)
                {
                    //TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    //TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    //TxtAPG1DaysRem.Text = tempdays.ToString();
                    TxtAPG1DaysRem.ForeColor = Color.Red;
                }
                //APG2
                if (TxtAPG2DaysRem.Text == "")
                {
                    TxtAPG2DaysRem.Text = 0.ToString();
                }
                tempdays = Convert.ToSingle(TxtAPG2DaysRem.Text);
                if (tempdays > 7)
                {
                    //TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    //TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    //TxtAPG2DaysRem.Text = tempdays.ToString();
                    TxtAPG2DaysRem.ForeColor = Color.Red;
                }
                //PB
                if (TxtPBDaysRem.Text == "")
                {
                    TxtPBDaysRem.Text = 0.ToString();
                }
                tempdays = Convert.ToSingle(TxtPBDaysRem.Text);
                if (tempdays > 7)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBDaysRem.ForeColor = Color.Red;
                }
                //Insurance
                if (TxtInsDaysRem.Text == "")
                {
                    TxtInsDaysRem.Text = 0.ToString();
                }
                tempdays = Convert.ToSingle(TxtInsDaysRem.Text);
                if (tempdays > 7)
                {
                    //TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    //TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    //TxtInsDaysRem.Text = tempdays.ToString();
                    TxtInsDaysRem.ForeColor = Color.Red;
                }
                //PB_FL
                if (TxtPBFLDaysRem.Text == "")
                {
                    TxtPBFLDaysRem.Text = 0.ToString();
                }
                tempdays = Convert.ToSingle(TxtPBFLDaysRem.Text);
                if (tempdays > 7)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.ForestGreen;
                }
                else if (tempdays <= 7 && tempdays > 0)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.Violet;
                }
                else if (tempdays <= 0)
                {
                    //TxtPBDaysRem.Text = tempdays.ToString();
                    TxtPBFLDaysRem.ForeColor = Color.Red;
                }

                string ProjectID = TxtProjectID.Text;

                string ContractID = TxtContractID.Text;
                string Ward = TxtWard.Text;
                string Location = TxtLocation.Text;

                TxtLog.AppendText("Displayed Projedt ID: " + ProjectID + " => " + Contract_ID + " of " + Ward + " at " + Location);
                TxtLog.AppendText(Environment.NewLine);
            }
        }

        private void modifyToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            string ProjectID = TxtProjectID.Text;
            string FiscalYear = TxtFY.Text;
            string ContractID = TxtContractID.Text;
            string ContractName = TxtContractName.Text;
            string ContractBudget = TxtContractBudget.Text;
            string Ward = TxtWard.Text;
            string ProjectType = TxtProjectType.Text;
            string BudgetType = TxtBudgetType.Text;
            string Location = TxtLocation.Text;

            string APG1DocRefNo = TxtAPG1RefNo.Text;
            string APG1Deadline = TxtAPG1DL.Text;
            string APG1Amount = TxtAPG1Amount.Text;
            string APG1MinDL = TxtAPG1MinDL.Text;
            string APG1Remark = TxtAPG1Remark.Text;

            string APG2DocRefNo = TxtAPG2RefNo.Text;
            string APG2Deadline = TxtAPG2DL.Text;
            string APG2Amount = TxtAPG2Amount.Text;
            string APG2MinDL = TxtAPG2MinDL.Text;
            string APG2Remark = TxtAPG2Remark.Text;

            string PBDocRefNo = TxtPBRefNo.Text;
            string PBDeadline = TxtPBDL.Text;
            string PBAmount = TxtPBAmount.Text;
            string PBMinDL = TxtPBMinDL.Text;
            string PBRemark = TxtPBRemark.Text;

            string InsDocRefNo = TxtInsRefNo.Text;
            string InsDeadline = TxtInsDL.Text;
            string InsAmount = TxtInsAmount.Text;
            string InsMinDL = TxtInsMinDL.Text;
            string InsRemark = TxtInsRemark.Text;

            string CurrentStatus = TxtCurrentStatus.Text;
            string NoticeIssued = TxtNoticeIssued.Text;
            string LOI = TxtLOI.Text;
            string LOA = TxtLOA.Text;
            string ContractAgreement = TxtContractAgreement.Text;
            string WorkPermit = TxtWorkPermit.Text;
            string WorkComplete = TxtWorkComplete.Text;
            string RunningBill = TxtRunningBill.Text;
            string FinalBill = TxtFinalBill.Text;
            string DaysRemaining = TxtDaysRem.Text;

            string NameOfContractor = TxtContractorName.Text;
            string AddressOfContractor = TxtAddressOfContractor.Text;
            string Email1 = TxtEmail1.Text;
            string ContractorOther = TxtContractorOther.Text;

            string ProjectDescription = TxtProjectDescription.Text;
            string Length = TxtLength.Text;
            string Breadth = TxtBreadth.Text;
            string Height = TxtHeight.Text;

            string ContractorNameDev = TxtContractorNameDev.Text;
            string ContractorAddressDev = TxtContractorAddressDev.Text;

            string APG1DaysRem = TxtAPG1DaysRem.Text;
            string APG2DaysRem = TxtAPG2DaysRem.Text;
            string PBDaysRem = TxtPBDaysRem.Text;
            string InsDaysRem = TxtInsDaysRem.Text;

            string APG1BankName = TxtBankNameAPG1.Text;
            string APG2BankName = TxtBankNameAPG2.Text;
            string PBBankName = TxtBankNamePB.Text;
            string InsBankName = TxtBankNameIns.Text;

            string APG1BankAddress = TxtBankAddressAPG1.Text;
            string APG2BankAddress = TxtBankAddressAPG2.Text;
            string PBBankAddress = TxtBankAddressPB.Text;
            string InsBankAddress = TxtBankAddressIns.Text;

            string ProcurementCategory = TxtProcurementcategory.Text;
            string ProcurementMethod = TxtProcurementMethod.Text;
            string TotalEstimatedAmount = TxtTotalEstimatedAmount.Text;
            string TotalContractAmount = TxtTotalContractAmount.Text;
            string TotalFinalBillAmount = TxtTotalFinalBillAmount.Text;
            string PublicEntity = TxtPE.Text;

            string PB2DocRefNo = TxtFL_PBRef.Text;
            string PB2Deadline = TxtFL_PBDeadline.Text;
            string PB2Amount = TxtFL_PBAmount.Text;
            string PB2MinDL = TxtFL_PBMinDL.Text;
            string PB2Remark = TxtFL_PBRemark.Text;
            string PB2BankName = TxtFL_BankNamePB.Text;
            string PB2BankAddress = TxtFL_BankAddressPB.Text;
            string PB2DaysRem = TxtPBFLDaysRem.Text;

            DialogResult dr = MessageBox.Show("Are you sure, you want to Modify?", "Modify", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                //Modify
                SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
                ConnectDb.Open();

                string query = "REPLACE INTO ContractTable(ProjectID,FiscalYear,ContractID,ContractName,ContractBudget,Ward," +
                    "ProjectType,BudgetType,Location,APG1DocRefNo,APG1Deadline, APG1Amount,APG1MinDL,APG1Remark," +
                    "APG2DocRefNo,APG2Deadline, APG2Amount,APG2MinDL,APG2Remark," +
                    "PBDocRefNo,PBDeadline, PBAmount,PBMinDL,PBRemark," +
                    "InsDocRefNo,InsDeadline, InsAmount,InsMinDL,InsRemark," +
                    "CurrentStatus,NoticeIssued,LOI,LOA,ContractAgreement,WorkPermit,WorkComplete,RunningBill,FinalBill,DaysRemaining," +
                    "NameOfContractor,AddressOfContractor,Email1,ContractorOther,ProjectDescription,Length,Breadth,Height,ContractorNameDev,ContractorAddressDev," +
                    "APG1DaysRem,APG2DaysRem,PBDaysRem,InsDaysRem,APG1BankName,APG2BankName ,PBBankName ,InsBankName,APG1BankAddress,APG2BankAddress,PBBankAddress,InsBankAddress," +
                    "ProcurementCategory, ProcurementMethod, TotalEstimatedAmount, TotalContractAmount, TotalFinalBillAmount, PublicEntity," +
                    "PB2DocRefNo,PB2Deadline, PB2Amount,PB2MinDL,PB2Remark,PB2BankName,PB2BankAddress,PB2DaysRem) " +
                    "VALUES('" + ProjectID + "', '" + FiscalYear + "','" + ContractID + "','" + ContractName + "','" + ContractBudget + "'," +
                    "'" + Ward + "','" + ProjectType + "','" + BudgetType + "','" + Location + "'" +
                    ",'" + APG1DocRefNo + "','" + APG1Deadline + "','" + APG1Amount + "','" + APG1MinDL + "','" + APG1Remark + "'" +
                    ",'" + APG2DocRefNo + "','" + APG2Deadline + "','" + APG2Amount + "','" + APG2MinDL + "','" + APG2Remark + "'" +
                    ",'" + PBDocRefNo + "','" + PBDeadline + "','" + PBAmount + "','" + PBMinDL + "','" + PBRemark + "'" +
                    ",'" + InsDocRefNo + "','" + InsDeadline + "','" + InsAmount + "','" + InsMinDL + "','" + InsRemark + "'" +
                    ",'" + CurrentStatus + "','" + NoticeIssued + "','" + LOI + "','" + LOA + "','" + ContractAgreement + "','" + WorkPermit + "'" +
                    ",'" + WorkComplete + "','" + RunningBill + "','" + FinalBill + "','" + DaysRemaining + "'" +
                    ",'" + NameOfContractor + "','" + AddressOfContractor + "','" + Email1 + "','" + ContractorOther + "'" +
                    ",'" + ProjectDescription + "','" + Length + "','" + Breadth + "','" + Height + "','" + ContractorNameDev + "','" + ContractorAddressDev + "'" +
                    ",'" + APG1DaysRem + "','" + APG2DaysRem + "','" + PBDaysRem + "','" + InsDaysRem + "'" +
                    ",'" + APG1BankName + "','" + APG2BankName + "','" + PBBankName + "','" + InsBankName + "'" +
                    ",'" + APG1BankAddress + "','" + APG2BankAddress + "','" + PBBankAddress + "','" + InsBankAddress + "'" +
                    ",'" + ProcurementCategory + "','" + ProcurementMethod + "','" + TotalEstimatedAmount + "', '" + TotalContractAmount + "', '" + TotalFinalBillAmount + "', '" + PublicEntity + "' " +
                    ", '" + PB2DocRefNo + "','" + PB2Deadline + "','" + PB2Amount + "','" + PB2MinDL + "','" + PB2Remark + "','" + PB2BankName + "','" + PB2BankAddress + "','" + PB2DaysRem + "')";// one data format  = '" + Height + "'

                SQLiteCommand Cmd = new SQLiteCommand(query, ConnectDb);
                Cmd.ExecuteNonQuery();

                ConnectDb.Close();

                //BtnCreateProjectFolder_Click(sender, e);
                createProjectFolderToolStripMenuItem_Click(sender, e);
                BtnSave2Txt_Click(sender, e);

                TxtLog.AppendText("Activity: Successfully Modified Record: " + "Project ID: " + ProjectID + "  " + ContractID + " of " + Ward + " at " + Location);
                TxtLog.AppendText(Environment.NewLine);

                /*using (System.IO.StreamWriter sw = System.IO.File.AppendText(@".\Log\Log.txt"))
                {
                    Text2Write = "[" + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "]" + "  --->  " + "MODIFY" + " ---> " + "Project ID: " + ProjectID + "  " + ProjectName + " of " + Ward + " at " + Location;
                    sw.WriteLine(Text2Write);
                }*/

            }
            else if (dr == DialogResult.No)
            {
                //Nothing to do
            }
        }

        private void deleteToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            string ProjectID = TxtProjectID.Text;

            if (TxtProjectID.Text == "")
            {
                TxtLog.Text = "Enter Project ID to Delete";
            }
            else
            {
                DialogResult dr = MessageBox.Show("Are You Sure, you want to delete?", "Delete", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    //delete
                    SQLiteConnection ConnectDb = new SQLiteConnection("Data Source = Contract.sqlite3");
                    ConnectDb.Open();

                    string query = "DELETE FROM  ContractTable WHERE ProjectID ='" + TxtProjectID.Text + "' ";
                    SQLiteCommand Cmd = new SQLiteCommand(query, ConnectDb);
                    Cmd.ExecuteNonQuery();

                    ConnectDb.Close();

                    TxtProjectID.Text = "";

                    string ContractID = TxtContractID.Text;
                    string Ward = TxtWard.Text;
                    string Location = TxtLocation.Text;
                    TxtLog.AppendText("Deleted Projedt ID: " + ProjectID + " => " + ContractID + " of " + Ward + " at " + Location);
                    TxtLog.AppendText(Environment.NewLine);

                    /*using (System.IO.StreamWriter sw = System.IO.File.AppendText(@".\Log\Log.txt"))
                    {
                        Text2Write = "[" + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "]" + "  --->  " + "DELETE" + " ---> " + "Project ID: " + ProjectID + "  " + ProjectName + " of " + Ward + " at " + Location;
                        sw.WriteLine(Text2Write);
                    }*/
                    DeleteTextFields();
                    Initial_State_of_Label();
                }
                else if (dr == DialogResult.No)
                {
                    //Nothing to do
                }

            }
        }

        private void createProjectFolderToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            Fun_CreateProjectFolder();
        }

        private void analyseDateToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            Fun_AnalyseDate();
        }

        private void createPdfToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            Fun_CreateAllPdf();
        }

        private void addToolStripMenuItem1_Click(global::System.Object sender, global::System.EventArgs e)
        {
            TxtProjectID.Enabled = false;
            TxtProjectID.Text = "";

            modifyToolStripMenuItem.Enabled = false;
            displayToolStripMenuItem.Enabled = false;
            deleteToolStripMenuItem.Enabled = false;
            addToolStripMenuItem.Enabled = true;
            //BtnModify.Enabled = false;
            //BtnDisplay.Enabled = false;
            //BtnDelete.Enabled = false;
            //BtnAdd.Enabled = true;

            addToolStripMenuItem1.Checked = true;
            displayModifyDeleteToolStripMenuItem.Checked = false;

            DeleteTextFields();
            BtnResetBill_Click(sender, e);
            Initial_State_of_Label();
        }

        private void displayModifyDeleteToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            TxtProjectID.Enabled = true;
            TxtProjectID.Text = "";

            modifyToolStripMenuItem.Enabled = true;
            displayToolStripMenuItem.Enabled = true;
            deleteToolStripMenuItem.Enabled = true;
            addToolStripMenuItem.Enabled = false;

            addToolStripMenuItem1.Checked = false;
            displayModifyDeleteToolStripMenuItem.Checked = true;

            DeleteTextFields();
            BtnResetBill_Click(sender, e);
            Initial_State_of_Label();
        }

        private void ComboBoxFormatBill_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {
            TxtFormatBill.Text = ComboBoxFormatBill.Text;
        }

        private void TxtProcurementcategory_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                //Add ---> Procurement Method
                ComboBoxProMethod.Items.Clear();
                string[] ProcMethodList = System.IO.File.ReadAllLines(@".\ComboBoxList\ProcurementMethod\" + TxtProcurementcategory.Text + ".txt");
                foreach (var line in ProcMethodList)
                {
                    ComboBoxProMethod.Items.Add(line);
                }
            }
            catch
            {

            }

        }

        private void ComboBoxProCategory_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {
            TxtProcurementcategory.Text = ComboBoxProCategory.Text;
        }

        private void ComboBoxProMethod_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {
            TxtProcurementMethod.Text = ComboBoxProMethod.Text;
        }

        private void BtnCalcDeduction_Click(global::System.Object sender, global::System.EventArgs e)
        {
            FrmCalcDeduction fcalcded = new FrmCalcDeduction();
            fcalcded.Show();
        }

        private void Est_Contract_Ratio()
        {
            try
            {
                double est, cont, ratio;
                est = Convert.ToDouble(TxtTotalEstimatedAmount.Text);
                cont = Convert.ToDouble(TxtTotalContractAmount.Text);

                ratio = Math.Round(cont / est * 100.0, 2);
                LblEst_Contract.Text = "Contract/Estimate ratio = " + ratio.ToString() + "%";

                if (ratio <= 100.0)
                {
                    LblEst_Contract.ForeColor = Color.ForestGreen;
                }
                else if (ratio > 100.0)
                {
                    LblEst_Contract.ForeColor = Color.Red;
                }
            }
            catch
            {

            }
        }

        private void Contract_Bill_Ratio()
        {
            try
            {
                double bill, cont, ratio;
                bill = Convert.ToDouble(TxtTotalFinalBillAmount.Text);
                cont = Convert.ToDouble(TxtTotalContractAmount.Text);

                ratio = Math.Round(bill / cont * 100.0, 2);
                LblBill_Contract.Text = "Bill/Contract ratio =  " + ratio.ToString() + "%";

                if (ratio <= 100.0)
                {
                    LblBill_Contract.ForeColor = Color.ForestGreen;
                }
                else if (ratio > 100.0)
                {
                    LblBill_Contract.ForeColor = Color.Red;
                }
            }
            catch
            {

            }
        }

        private void TxtTotalEstimatedAmount_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                Est_Contract_Ratio();
            }
            catch
            {

            }
        }

        private void TxtTotalContractAmount_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                Est_Contract_Ratio();
                Contract_Bill_Ratio();
            }
            catch
            {

            }
        }

        private void TxtTotalFinalBillAmount_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                Contract_Bill_Ratio();
            }
            catch
            {

            }
        }

        private void BtnCheckPB_Click(global::System.Object sender, global::System.EventArgs e)
        {
            FrmPBAmount fpb = new FrmPBAmount();
            fpb.Show();
        }

        private void TxtContractID_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                bool DoIDExists = IsContractIDUnique();
                if (DoIDExists == true)
                {
                    LblUniqueID.Text = "Already in Use";
                    LblUniqueID.ForeColor = Color.Red;
                    TxtContractID.BackColor = Color.LightSalmon;
                }
                else
                {
                    LblUniqueID.Text = "Available";
                    LblUniqueID.ForeColor = Color.ForestGreen;
                    TxtContractID.BackColor = Color.LightGreen;
                }

                if (TxtContractID.Text == "")
                {
                    LblUniqueID.Text = "Available/Already in Use";
                    LblUniqueID.ForeColor = Color.Black;
                    TxtContractID.BackColor = Color.White;
                }
            }
            catch
            {

            }
        }

        private void TxtWorkPermit_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                dataGridView3.Rows[1].Cells[1].Value = TxtWorkPermit.Text;
            }
            catch
            {

            }
        }

        private void TxtContractAgreement_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                dataGridView3.Rows[0].Cells[1].Value = TxtContractAgreement.Text;
            }
            catch
            {

            }
        }

        private void TxtWorkComplete_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                dataGridView3.Rows[2].Cells[1].Value = TxtWorkComplete.Text;
            }
            catch
            {

            }
        }

        private void BtnWCToOld_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                dataGridView3.Rows[3].Cells[1].Value = TxtWorkComplete.Text;
            }
            catch
            {

            }
        }

        private void workPermitLetterToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            if (TxtFY.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "" || TxtContractID.Text == "")
            {
                TxtLog.Text = "Please fill mandator fields to continue !!!";
                //TxtLog.Text += Environment.NewLine;
            }
            else
            {
                string ThisDir = Environment.CurrentDirectory;
                //string FontDir1 = ThisDir + "\\Font\\Preeti Normal.otf";
                // path folder
                CreateAccessProjectFolders();
                string filename_docx = EventHistoryFolder + "\\WorkPermitLetter.docx";

                //CreateAccessProjectFolders();

                //Start Word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;
                oWord = new Word.Application();
                oWord.Visible = false;

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                Cur_Dir = Environment.CurrentDirectory;
                string filename_template = Cur_Dir + "\\ComboBoxList\\LetterFormat\\WorkPermit_Template.dotx";
                object oTemplate = filename_template;
                //object oTemplate = "E:\\Tippani_Template.dotx";

                oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

                //Bookmarks and Data
                object oBookMark;
                oBookMark = "WorkPermitDate_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[1].Cells[2].Value.ToString();

                oBookMark = "ContractorName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorNameDev.Text;

                oBookMark = "ContractorAddress_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorAddressDev.Text;

                oBookMark = "ContractDateBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[0].Cells[2].Value.ToString();

                oBookMark = "ContractDateAD_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[0].Cells[1].Value.ToString();

                oBookMark = "ContractName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractName.Text;

                oBookMark = "ContractID_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractID.Text;

                oBookMark = "WorkCompletionDateBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[2].Cells[2].Value.ToString();

                oBookMark = "WorkCompletionDateAD_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[2].Cells[1].Value.ToString();

                //string filename_docx = Cur_Dir + "\\InputFolder\\NewLetter.docx"; 
                //string filename_docx = Project_Folders + "\\" + TxtFirstName.Text + "_" + TxtPlotNo.Text + "_Letter.docx";

                oDoc.SaveAs2(filename_docx);

                oDoc.Close();
                oWord.Quit();


                //TxtRecentFolderLocation.Text = Project_Folders;
                TxtLog.Text = "Work Permit letter Saved.";
            }
        }

        private void extensionOfTimeLetterToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            if (TxtFY.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "" || TxtContractID.Text == "")
            {
                TxtLog.Text = "Please fill mandator fields to continue !!!";
                //TxtLog.Text += Environment.NewLine;
            }
            else
            {
                string ThisDir = Environment.CurrentDirectory;
                //string FontDir1 = ThisDir + "\\Font\\Preeti Normal.otf";
                // path folder
                CreateAccessProjectFolders();
                string filename_docx = EventHistoryFolder + "\\EOT.docx";

                //CreateAccessProjectFolders();


                //Start Word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;
                oWord = new Word.Application();
                oWord.Visible = false;

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                Cur_Dir = Environment.CurrentDirectory;
                string filename_template = Cur_Dir + "\\ComboBoxList\\LetterFormat\\EOT_Template.dotx";
                object oTemplate = filename_template;
                //object oTemplate = "E:\\Tippani_Template.dotx";

                oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

                //Bookmarks and Data
                object oBookMark;
                oBookMark = "EOTDate_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[4].Cells[2].Value.ToString();

                oBookMark = "ContractorName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorNameDev.Text;

                oBookMark = "ContractorAddressBM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorAddressDev.Text;

                oBookMark = "ContractName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractName.Text;

                oBookMark = "ContractID_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractID.Text;

                oBookMark = "EOT_Time_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = LblEOT.Text;

                oBookMark = "WorkCompletionDateOldBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[2].Cells[2].Value.ToString();

                oBookMark = "WorkCompletionDateNewBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[3].Cells[2].Value.ToString();

                //string filename_docx = Cur_Dir + "\\InputFolder\\NewLetter.docx"; 
                //string filename_docx = Project_Folders + "\\" + TxtFirstName.Text + "_" + TxtPlotNo.Text + "_Letter.docx";

                oDoc.SaveAs2(filename_docx);

                oDoc.Close();
                oWord.Quit();


                //TxtRecentFolderLocation.Text = Project_Folders;
                TxtLog.Text = "EOT letter Saved.";
            }
        }

        private void CreateEOT()
        {
            try
            {
                string y, m, d;
                string yn, mn, dn;
                yn = " वर्ष ";
                mn = " महिना ";
                dn = " दिन ";
                y = TxtEOTYear.Text;
                m = TxtEOTMonth.Text;
                d = TxtEOTDay.Text;
                if (y == "0" || y == "")
                {
                    yn = "";
                }
                if (m == "0" || m == "")
                {
                    mn = "";
                }
                if (d == "0" || d == "")
                {
                    dn = "";
                }
                LblEOT.Text = y + yn + m + mn + d + dn;
            }
            catch
            {

            }
        }

        private void TxtEOTYear_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            CreateEOT();
        }

        private void TxtEOTMonth_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            CreateEOT();
        }

        private void TxtEOTDay_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            CreateEOT();
        }

        private void BtnFindNewWCdate_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                int days, mon, yr;
                days = Convert.ToInt32(TxtEOTDay.Text);
                mon = Convert.ToInt32(TxtEOTMonth.Text);
                yr = Convert.ToInt32(TxtEOTYear.Text);

                mon += yr * 12;
                string Olddates;
                Olddates = dataGridView3.Rows[3].Cells[1].Value.ToString();
                TxtWorkComplete.Text = NewDateAFterAddingDays_and_Months(days, mon, Olddates);
            }
            catch
            {

            }

        }


        private void BtnFindEOT_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                //string Olddates, newdates;

                string Olddates = dataGridView3.Rows[3].Cells[1].Value.ToString();
                string newdates = dataGridView3.Rows[2].Cells[1].Value.ToString();

                int year1, month1, days1, year2, month2, days2;
                string[] temp_date1 = Olddates.Split("-");
                string[] temp_date2 = newdates.Split("-");

                /*int[] monthdays = new int[]
                {
                    31,28,31,30,31,30,31,31,30,31,30,31
                };*/

                year1 = Convert.ToInt32(temp_date1[0]);
                month1 = Convert.ToInt32(temp_date1[1]);
                days1 = Convert.ToInt32(temp_date1[2]);

                year2 = Convert.ToInt32(temp_date2[0]);
                month2 = Convert.ToInt32(temp_date2[1]);
                days2 = Convert.ToInt32(temp_date2[2]);

                LocalDate start = new LocalDate(year1, month1, days1);
                LocalDate end = new LocalDate(year2, month2, days2);



                /*int[] EOT_Dur = DifferenceInDateYYMMDD(Olddates, newdates);
                TxtEOTYear.Text = EOT_Dur[0].ToString();
                TxtEOTMonth.Text = EOT_Dur[1].ToString();
                TxtEOTDay.Text = EOT_Dur[2].ToString();
                label86.Text = EOT_Dur.ToString();*/

                Period agePeriod = Period.Between(start, end, PeriodUnits.YearMonthDay);

                TxtEOTYear.Text = agePeriod.Years.ToString();
                TxtEOTMonth.Text = agePeriod.Months.ToString();
                TxtEOTDay.Text = agePeriod.Days.ToString();

            }
            catch
            {

            }

        }

        private void tippaniForAdvancePayment1ToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            if (TxtFY.Text == "" || TxtWard.Text == "" || TxtProjectType.Text == "" || TxtContractID.Text == "")
            {
                TxtLog.Text = "Please fill mandatory fields to continue !!!";
                //TxtLog.Text += Environment.NewLine;
            }
            else
            {
                string ThisDir = Environment.CurrentDirectory;
                //string FontDir1 = ThisDir + "\\Font\\Preeti Normal.otf";
                // path folder
                CreateAccessProjectFolders();
                string filename_docx = EventHistoryFolder + "\\Tippani_AGP1.docx";

                //CreateAccessProjectFolders();


                //Start Word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;
                oWord = new Word.Application();
                oWord.Visible = false;

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                Cur_Dir = Environment.CurrentDirectory;
                string filename_template = Cur_Dir + "\\ComboBoxList\\LetterFormat\\Tippani_AGP1_Template.dotx";
                object oTemplate = filename_template;
                //object oTemplate = "E:\\Tippani_Template.dotx";

                oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

                //Bookmarks and Data
                object oBookMark;
                oBookMark = "TippaniDate_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView4.Rows[2].Cells[1].Value.ToString();

                oBookMark = "ContractorName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorNameDev.Text;

                oBookMark = "ContractorAddressBM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractorAddressDev.Text;

                oBookMark = "ContractDateBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[0].Cells[2].Value.ToString();

                oBookMark = "ContractDateAD_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[0].Cells[1].Value.ToString();

                oBookMark = "ContractName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractName.Text;

                oBookMark = "ContractID_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtContractID.Text;

                oBookMark = "ReqLetterDateBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView4.Rows[0].Cells[1].Value.ToString();

                oBookMark = "AP_Percent_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView4.Rows[1].Cells[1].Value.ToString();

                oBookMark = "APGIssueDateBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[5].Cells[2].Value.ToString();

                oBookMark = "APGIssueDateAD_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[5].Cells[1].Value.ToString();

                oBookMark = "APGBankName_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtBankNameAPG1.Text;

                oBookMark = "APGDLBS_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[7].Cells[2].Value.ToString();

                oBookMark = "APGDLAD_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView3.Rows[8].Cells[1].Value.ToString();

                oBookMark = "APGAmount_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView4.Rows[3].Cells[1].Value.ToString();

                oBookMark = "APGRef_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = TxtAPG1RefNo.Text;

                oBookMark = "ContractPriceST_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView1.Rows[1].Cells[3].Value.ToString();

                oBookMark = "AP_Percent1_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView4.Rows[1].Cells[1].Value.ToString();

                oBookMark = "AP_Amount_BM";
                oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = dataGridView1.Rows[9].Cells[3].Value.ToString();

                //string filename_docx = Cur_Dir + "\\InputFolder\\NewLetter.docx"; 
                //string filename_docx = Project_Folders + "\\" + TxtFirstName.Text + "_" + TxtPlotNo.Text + "_Letter.docx";

                oDoc.SaveAs2(filename_docx);

                oDoc.Close();
                oWord.Quit();


                //TxtRecentFolderLocation.Text = Project_Folders;
                TxtLog.Text = "Tippani for AP-1 Saved.";
            }
        }

        private void TxtAPG1DL_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            dataGridView3.Rows[7].Cells[1].Value = TxtAPG1DL.Text;
        }

        private void TxtAPG2DL_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            dataGridView3.Rows[8].Cells[1].Value = TxtAPG2DL.Text;
        }

        private void TxtAPG1Amount_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            dataGridView4.Rows[3].Cells[1].Value = TxtAPG1Amount.Text;
        }

        private void TxtAPG2Amount_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            dataGridView4.Rows[3].Cells[2].Value = TxtAPG2Amount.Text;
        }

        private void convertADToBSToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                //load default format bill format text name
                string[] BaseDatesAll = System.IO.File.ReadAllLines(@".\ComboBoxList\ADBSConversion\BaseDate.txt");
                string BaseDateAD = BaseDatesAll[0].Split('\t')[1];
                string BaseDateBS = BaseDatesAll[1].Split('\t')[1];
                int BaseYearBS = Convert.ToInt32(BaseDateBS.Split('-')[0]);

                //MessageBox.Show("BS = " + BaseDateBS + "\nAD = " + BaseDateAD);
                int No_of_Month_in_List = 0, thisyearBS;
                string appendedDays = "";
                //string[] days_of_each_Month_in_List = new string[120];
                string directory1 = Environment.CurrentDirectory + "\\ComboBoxList\\ADBSConversion\\DaysEachMonth";

                //load format bill in combobox
                //string dir = Environment.CurrentDirectory + "\\ComboBoxList\\BillFormat";
                string[] files = Directory.GetFiles(directory1, "*.txt", SearchOption.TopDirectoryOnly);//Directory.GetFiles(dir);
                int countyearfiles = 0;
                foreach (string filePath in files)
                {
                    int yearfilename = Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(filePath));
                    if (yearfilename >= BaseYearBS)
                    {
                        countyearfiles++;
                    }

                    //int fileCount = Directory.GetFiles(directory1, "*.txt", SearchOption.TopDirectoryOnly).Length;
                }

                int no_of_Yrs = countyearfiles;
                for (int i = 0; i < no_of_Yrs; i++)
                {
                    thisyearBS = BaseYearBS + i;
                    string filename = directory1 + "\\" + thisyearBS.ToString() + ".txt";
                    string[] EachDaysinMonth = System.IO.File.ReadAllLines(filename);
                    int No_of_Month_in_Year1 = EachDaysinMonth[0].Split(',').Length;

                    No_of_Month_in_List += No_of_Month_in_Year1;
                    if (i != (no_of_Yrs - 1)) appendedDays += EachDaysinMonth[0] + ',';
                    else appendedDays += EachDaysinMonth[0];
                    //days_of_each_Month_in_List = EachDaysinMonth[0].Split(',');
                }

                string[] days_of_each_Month_in_List = appendedDays.Split(',');
                //MessageBox.Show("appendeddays = " + appendedDays);
                //MessageBox.Show("months = " + No_of_Month_in_List + "\nDays = " + days_of_each_Month_in_List[7]);

                CSAYADtoBSConverter AD2BS = new CSAYADtoBSConverter();

                int countrows = dataGridView3.Rows.Count - 1;
                for (int i = 0; i < countrows; i++)
                {
                    string thisDateAD = "";
                    if (dataGridView3.Rows[i].Cells[1].Value != null)
                    {
                        //MessageBox.Show("row = " + i.ToString() + "\nDBVAL = " + dataGridView3.Rows[i].Cells[1].Value);
                        thisDateAD = dataGridView3.Rows[i].Cells[1].Value.ToString();
                        //CSAYADtoBSConverter AD2BS = new CSAYADtoBSConverter();
                        int daysdiff = AD2BS.DifferenceInDate(BaseDateAD, thisDateAD);

                        string newdateBS = AD2BS.Add_days_to_BS_Date(BaseDateBS, daysdiff, No_of_Month_in_List, days_of_each_Month_in_List);
                        dataGridView3.Rows[i].Cells[2].Value = newdateBS;
                    }

                }
            }
            catch
            {

            }


        }

        private void dateValidationToolStripMenuItem_Click(global::System.Object sender, global::System.EventArgs e)
        {
            string infoDate = "";
            //load default format bill format text name
            string[] BaseDatesAll = System.IO.File.ReadAllLines(@".\ComboBoxList\ADBSConversion\BaseDate.txt");
            string BaseDateAD = BaseDatesAll[0].Split('\t')[1];
            string BaseDateBS = BaseDatesAll[1].Split('\t')[1];
            int BaseYearBS = Convert.ToInt32(BaseDateBS.Split('-')[0]);

            infoDate += "TIME PERIOD WITHIN WHICH DATE CAN BE CONVERTED\n";
            infoDate += "------------------------------------------------------------------------\n";
            infoDate += "\nBase Date AD = " + BaseDateAD.ToString();
            infoDate += "\nBase Date BS = " + BaseDateBS.ToString();

            //MessageBox.Show("BS = " + BaseDateBS + "\nAD = " + BaseDateAD);
            int No_of_Month_in_List = 0, thisyearBS;
            string appendedDays = "";
            //string[] days_of_each_Month_in_List = new string[120];
            string directory1 = Environment.CurrentDirectory + "\\ComboBoxList\\ADBSConversion\\DaysEachMonth";

            //load format bill in combobox
            //string dir = Environment.CurrentDirectory + "\\ComboBoxList\\BillFormat";
            string[] files = Directory.GetFiles(directory1, "*.txt", SearchOption.TopDirectoryOnly);//Directory.GetFiles(dir);
            int countyearfiles = 0;
            foreach (string filePath in files)
            {
                int yearfilename = Convert.ToInt32(System.IO.Path.GetFileNameWithoutExtension(filePath));
                if (yearfilename >= BaseYearBS)
                {
                    countyearfiles++;
                }

                //int fileCount = Directory.GetFiles(directory1, "*.txt", SearchOption.TopDirectoryOnly).Length;
            }

            int no_of_Yrs = countyearfiles;
            for (int i = 0; i < no_of_Yrs; i++)
            {
                thisyearBS = BaseYearBS + i;
                string filename = directory1 + "\\" + thisyearBS.ToString() + ".txt";
                string[] EachDaysinMonth = System.IO.File.ReadAllLines(filename);
                int No_of_Month_in_Year1 = EachDaysinMonth[0].Split(',').Length;

                No_of_Month_in_List += No_of_Month_in_Year1;
                if (i != (no_of_Yrs - 1)) appendedDays += EachDaysinMonth[0] + ',';
                else appendedDays += EachDaysinMonth[0];
                //days_of_each_Month_in_List = EachDaysinMonth[0].Split(',');
            }

            infoDate += "\n\nNo. of months starting from base date = " + No_of_Month_in_List.ToString();

            string[] days_of_each_Month_in_List = appendedDays.Split(',');

            //summation of total days starting from base date
            int[] IntDays = new int[No_of_Month_in_List];
            int sum = 0;
            for (int i = 0; i < No_of_Month_in_List; i++)
            {
                IntDays[i] = Convert.ToInt32(days_of_each_Month_in_List[i]);
                sum += IntDays[i];
            }

            infoDate += "\nNo. of days starting from base date = " + sum.ToString();

            int days = sum;
            string oldDate = BaseDateAD;
            string LastDateAD = NewDateAFterAddingDays_and_Months(days, 0, oldDate);
            infoDate += "\n\nLast date AD up to which date can be converted = " + LastDateAD;

            //oldDate = BaseDateBS;
            int days2add = sum;
            CSAYADtoBSConverter ad2bs = new CSAYADtoBSConverter();
            string LastDateBS = ad2bs.Add_days_to_BS_Date(BaseDateBS, days2add, No_of_Month_in_List, days_of_each_Month_in_List);
            infoDate += "\nLast date BS up to which date can be converted = " + LastDateBS;
            infoDate += "\n\nNote: It can only convert date between Basedate and Lastdate as mentioned above";
            infoDate += "\n\nNote: To convert beyond the above bound, you can add days of months to .txt or/and change base date.";
            MessageBox.Show(infoDate);


        }
    }
}
