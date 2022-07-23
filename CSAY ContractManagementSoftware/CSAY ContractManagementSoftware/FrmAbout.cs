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
    public partial class FrmAbout : Form
    {
        public FrmAbout()
        {
            InitializeComponent();
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmAbout_Load(object sender, EventArgs e)
        {
            textBox1.TabStop = false;
            textBox2.TabStop = false;
            textBox3.TabStop = false;


        }
    }
}
