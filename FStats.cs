using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sklad
{
    public partial class FStats : Form
    {
        public FStats()
        {
            InitializeComponent();
        }

        public void SetDataTable(DataTable dt)
        {
            dgStats.DataSource = dt;
            dgStats.AutoResizeColumns();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
