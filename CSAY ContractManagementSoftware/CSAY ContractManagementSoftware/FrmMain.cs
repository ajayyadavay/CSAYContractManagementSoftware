namespace CSAY_ContractManagementSoftware
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void BtnContract_Click(object sender, EventArgs e)
        { 
            FrmContract fcontract = new FrmContract();
            fcontract.Show();
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}