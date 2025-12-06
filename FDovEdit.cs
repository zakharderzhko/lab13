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
    public partial class FDovEdit : Form
    {
        public FDovEdit()
        {
            InitializeComponent();
            dgGrupa.DataSource = Form1.MySkladStatic.DovGrupa;
            dgPostachalnyk.DataSource = Form1.MySkladStatic.DovPostachalnyk;
            dgOdiVym.DataSource = Form1.MySkladStatic.DovOdiVym;
            dgSklady.DataSource = Form1.MySkladStatic.DovSklady;

            dgGrupa.Columns["Група"].HeaderText = "Група товарів";
            dgPostachalnyk.Columns["Постачальник"].HeaderText = "Постачальник";
            dgOdiVym.Columns["Одиниця"].HeaderText = "Одиниця виміру";
            dgSklady.Columns["Склад"].HeaderText = "Склад";

            dgGrupa.ReadOnly = false;
            dgPostachalnyk.ReadOnly = false;
            dgOdiVym.ReadOnly = false;
            dgSklady.ReadOnly = false;

            dgGrupa.AllowUserToAddRows = true;
            dgGrupa.AllowUserToDeleteRows = true;
            dgPostachalnyk.AllowUserToAddRows = true;
            dgPostachalnyk.AllowUserToDeleteRows = true;
            dgOdiVym.AllowUserToAddRows = true;
            dgOdiVym.AllowUserToDeleteRows = true;
            dgSklady.AllowUserToAddRows = true;
            dgSklady.AllowUserToDeleteRows = true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
