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
    public partial class FServ : Form
    {
        public FServ()
        {
            InitializeComponent();
            FServTB.Text = Form1.GlStringParameter; // Якщо фільтр уже був, то відразу записуємо його у вікно введення
        }

        private void FservBOk_Click(object sender, EventArgs e)
        {
            Form1.GlStringParameter = FServTB.Text; // записуємо введений критерій у основну форму
            Close(); // закриваємо форму Fserv, щоб далі продовжувалось виконання програми
        }
    }
}
