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
    public partial class Form1 : Form
    {
        public static TSklad MySkladStatic;
        public Form1()
        {
            InitializeComponent();
            this.DGSklad.CellValueChanged += new 
                System.Windows.Forms.DataGridViewCellEventHandler(this.DGSklad_CellValueChanged);
        }

        private TSklad MySklad;

        public static string GlStringParameter;


        private void Form1_Load(object sender, EventArgs e)
        {
            MySklad = new TSklad();

            // Пов'язуємо DataGridView, що на формі Form1 із SkladView, який поставлено на таблицю TabSklad:
            DGSklad.DataSource = MySklad.SkladView;

            MySklad.CreateDovGrupa(); // Створимо довідник груп
            // Зробимо для групи товару стовпчик типу DataGridViewComboBoxColumn
            MySklad.AddComboGrupa(DGSklad);

            MySklad.AddComboPostachalnyk(DGSklad);
            MySklad.AddComboOdiVym(DGSklad);

            // Створимо список елементів комбобоксу із рядків таблиця довідника груп
            foreach (DataRow r in MySklad.DovGrupa.Rows) // Для кожного рядка rr із таблиці DovGrupa
            {
                string s = (string)r["Група"];
                CBGrupa.Items.Add(r["Група"]); // Додаємо у "випадайку" контролза ComboBox елементи із довідника
            }

            foreach (DataRow r in MySklad.DovPostachalnyk.Rows) // Для кожного рядка rr із таблиці DovGrupa
            {
                string s = (string)r["Постачальник"];
                CBPostachalnyk.Items.Add(r["Постачальник"]); // Додаємо у "випадайку" контролза ComboBox елементи із довідника
            }

            foreach (DataRow r in MySklad.DovOdiVym.Rows) // Для кожного рядка rr із таблиці DovGrupa
            {
                string s = (string)r["Одиниця"];
                CBOdiVym.Items.Add(r["Одиниця"]); // Додаємо у "випадайку" контролза ComboBox елементи із довідника
            }
            MySkladStatic = MySklad;
            DGSklad.Columns["Одиниця"].DisplayIndex = 7;
            DGSklad.Columns["Кількість"].DisplayIndex = 6;

            foreach (DataRow r in MySklad.DovSklady.Rows)
                CBSklad.Items.Add(r["Склад"]);

            tvSklady.Nodes.Clear();
            foreach (DataRow rs in MySklad.DovSklady.Rows)
            {
                string skladName = rs["Склад"].ToString();
                TreeNode skladNode = new TreeNode(skladName);
                foreach (DataRow rg in MySklad.DovGrupa.Rows)
                {
                    TreeNode groupNode = new TreeNode(rg["Група"].ToString());
                    skladNode.Nodes.Add(groupNode);
                }
                tvSklady.Nodes.Add(skladNode);
            }
            tvSklady.ExpandAll();

            tvSklady.AfterSelect += TvSklady_AfterSelect;


            MySklad.ColumnPropSet(DGSklad);
            MySklad.SetSumy(DGSkladSum);

        }

        private void BAddRowToTable_Click(object sender, EventArgs e)
        {
            int pKilkist;
            decimal pPcina;

            // Зчитуємо значення
            if (!int.TryParse(TBKilkist.Text, out pKilkist))
            {
                MessageBox.Show("Введіть коректну кількість");
                return;
            }

            if (!decimal.TryParse(TBCina.Text, out pPcina))
            {
                MessageBox.Show("Введіть коректну ціну");
                return;
            }

            // Додаємо рядок
            MySklad.TSkladAddRow(
                CBGrupa.Text,
                CBSklad.Text,
                TBNazva.Text,
                TBVyrobnyk.Text,
                CBPostachalnyk.Text,
                pKilkist,
                pPcina,
                CBOdiVym.Text
            );

            // Оновлюємо підсумки
            MySklad.SetSumy(DGSkladSum);
        }

        private void записатиТаблицюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MySklad.ZapTabFile();
            MessageBox.Show("Таблиця записана");
        }

        private void зчитатиТаблицюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MySklad.ReadTabFile(DGSkladSum); // зчитати таблицю і сформувати підсумки у гріді DGSkladSum
            DGSklad.DataSource = MySklad.SkladView;
            RebuildTreeView();
        }

        private void DGSklad_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int i, j; decimal vart, kilk, cin; i = e.RowIndex; // індекс рядка
            j = e.ColumnIndex; // індекс стовпця
            if (i < 0) return; // якщо i < 0 або j < 0, то це – заголовок стовпця або рядка
            if (j < 0) return;
            if ((DGSklad.Columns[j].Name == "Кількість") || (DGSklad.Columns[j].Name == "Ціна")) // Якщо змінювалась ціна або кількість
            {
                try // Спробуємо, бо, можливо, введено не числа у поле ціни або кількості
                {
                    cin = (decimal)DGSklad.Rows[i].Cells["Ціна"].Value;
                    kilk = Convert.ToDecimal((Int32)DGSklad.Rows[i].Cells["Кількість"].Value);
                    vart = kilk * cin; DGSklad.Rows[i].Cells["Вартість"].Value = vart; // записуємо вартість у комірку гріда
                }
                catch { }
            }
            MySklad.SetSumy(DGSkladSum); // оновлюємо підсумки
        }

        private void встановитиФільтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form FiltrDialog = new FServ(); // створимо екземпляр форми FServ і назвемо його FiltrDialog // Встановлюємо текст заголовку форми FServ
            FiltrDialog.Text = "Введіть критерій фільтруванна - наприклад: Група = 'Книги' & Ціна < 70";
            GlStringParameter = MySklad.FiltrCriteria; /* Відкриємо форму у режимі діалогу. Це означає, що наступний оператор, що слідує за оператором FiltrDialog.ShowDialog(); буде виконуватись лише після того, як буде закрито форму, яку викликали */
            FiltrDialog.ShowDialog(); MySklad.TSkladValFiltr(GlStringParameter, DGSklad); // виклик методу встановлення фільтру
        }

        private void знятиФільтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlStringParameter = ""; // відмінити фільтрування таблиці
            MySklad.TSkladValFiltr(GlStringParameter, DGSklad);
        }

        private void встановитиКритерійСортуванняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form SortDialog = new FServ(); // створюємо екземпляр форми FServ і називаємо його SortDialog
                                           // Встановимо текст заголовку форми FServ
            SortDialog.Text = "Введіть критерій сортування - наприклад: Виробник, Ціна Desc";
            GlStringParameter = MySklad.SortCriteria;
            /* Відкриємо форму у режимі діалогу. Це означає, що оператор, який слідує за оператором 
             FiltrDialog.ShowDialog(); буде виконуватись лише після того, як буде закрито форму, 
             яку було викликано (тут – SortDialog) */
            SortDialog.ShowDialog();
            MySklad.TSkladValSort(GlStringParameter, DGSklad, DGSkladSum);
        }

        private void сортуватиПоГрупіToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GlStringParameter = "Група, Назва";
            MySklad.TSkladValSort(GlStringParameter, DGSklad, DGSkladSum);
        }

        private void пошукПоНазвіToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sNazva; // встановимо текст заголовку форми FServ
            Form SeekDialog = new FServ(); // створимо екземпляр форми FServ і назвемо його SeekDialog
            SeekDialog.Text = "Введіть назву:";
            /* Відкриємо форму у режимі діалогу. Це означає, що оператор, який слідує за оператором SeekDialog.ShowDialog(); буде виконуватись лише після того, коли буде закрито форму, яку було викликано (тут SeekDialog) */
            SeekDialog.ShowDialog();
            MySklad.SeekNazva(GlStringParameter, DGSklad);
        }

        private void DGSklad_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            /* Тут е – екземпляр класу DataGridViewCellValidatingEventArgs, аргумент, який передає значення індексів комірки, у яку вводили дані.
            e.rowindex – індекс рядка;
            e.columnindex – індекс стовпця;
            e.FormattedValue – значення введених у комірку даних */
            decimal cin; Int32 kilk;
            if (DGSklad.Columns[e.ColumnIndex].Name == "Ціна") // Якщо змінювалась ціна
            {
                if (DGSklad.Rows[e.RowIndex].IsNewRow)
                {
                    return;
                } // Не можна перевіряти дані у новому рядку
                  // Перевіряємо, чи введені дані можна трактувати як десяткове число
                if (!decimal.TryParse(e.FormattedValue.ToString(), out cin))
                // Метод TryParse поверне значення true, якщо можна і false, якщо ні. e.FormattedValue – введене значення
                {
                    MessageBox.Show("Введіть, будь ласка, числове значення у поле ціни.");
                    // e.Cancel = true; - Відмінити введення не правильного значення
                }
            }
            if (DGSklad.Columns[e.ColumnIndex].Name == "Кількість") // Якщо змінювалась кількість
            {
                if (DGSklad.Rows[e.RowIndex].IsNewRow) { return; } // Не можна перевіряти дані у новому рядку
                                                                   // Перевіряємо, чи введені дані можна трактувати як ціле число
                                                                   // Метод TryParse поверне значення true, якщо можна і false, якщо ні. e.FormattedValue – введене значення
                if (!Int32.TryParse(e.FormattedValue.ToString(), out kilk))
                {
                    MessageBox.Show("Введіть, будь ласка, цілочислове значення у поле кількості.");
                    e.Cancel = true; // Відмінити введення не правильного значення
                }
            }
        }

        private void редагуватиДовідникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (FDovEdit f = new FDovEdit())
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    CBGrupa.Items.Clear();
                    CBPostachalnyk.Items.Clear();
                    CBOdiVym.Items.Clear();
                    CBSklad.Items.Clear();

                    foreach (DataRow r in MySklad.DovGrupa.Rows)
                        CBGrupa.Items.Add(r["Група"]);

                    foreach (DataRow r in MySklad.DovPostachalnyk.Rows)
                        CBPostachalnyk.Items.Add(r["Постачальник"]);

                    foreach (DataRow r in MySklad.DovOdiVym.Rows)
                        CBOdiVym.Items.Add(r["Одиниця"]);

                    foreach (DataRow r in MySklad.DovSklady.Rows)
                        CBSklad.Items.Add(r["Склад"]);

                    MySklad.UpdateAllCombos(DGSklad);

                    MessageBox.Show("Довідники успішно оновлено!");
                }
            }
        }

        private void TvSklady_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // Якщо вибрано вузол складу — показати всі товари на вибраному складі
            // Якщо вибрано підвузол (група) — показати фільтр по складу+групі
            string filter = "";
            if (e.Node.Parent == null)
            {
                // вибрано склад
                string sklad = e.Node.Text.Replace("'", "''");
                filter = $"Склад = '{sklad}'";
            }
            else
            {
                string sklad = e.Node.Parent.Text.Replace("'", "''");
                string grupa = e.Node.Text.Replace("'", "''");
                filter = $"Склад = '{sklad}' AND Група = '{grupa}'";
            }

            MySklad.TSkladValFiltr(filter, DGSklad);
            MySklad.SetSumy(DGSkladSum);
        }

        private void RebuildTreeView()
        {
            tvSklady.Nodes.Clear();
            foreach (DataRow rs in MySklad.DovSklady.Rows)
            {
                TreeNode skladNode = new TreeNode(rs["Склад"].ToString());
                foreach (DataRow rg in MySklad.DovGrupa.Rows)
                    skladNode.Nodes.Add(new TreeNode(rg["Група"].ToString()));
                tvSklady.Nodes.Add(skladNode);
            }
            tvSklady.ExpandAll();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // Простий текстовий друк — виведемо заголовки та перші N рядків
            float y = 20;
            Font font = new Font("Arial", 10);
            // Заголовок
            e.Graphics.DrawString("Звіт по складах", new Font("Arial", 14, FontStyle.Bold), Brushes.Black, 20, y);
            y += 30;

            // Друк колонок
            string header = "№  Група    Склад    Назва    К-сть    Ціна    Вартість";
            e.Graphics.DrawString(header, font, Brushes.Black, 20, y);
            y += 20;

            // Друкуємо по TabSklad (обмежимо, наприклад, 100 рядків)
            int printed = 0;
            foreach (DataRow r in MySklad.TabSklad.Rows)
            {
                string line = $"{r["N_пп"],-3} {r["Група"],-10} {r["Склад"],-10} {r["Назва"],-18} {r["Кількість"],5} {r["Ціна"],10} {r["Вартість"],12}";
                e.Graphics.DrawString(line, font, Brushes.Black, 20, y);
                y += 18;
                printed++;
                if (y > e.MarginBounds.Bottom - 50 || printed > 100)
                {
                    e.HasMorePages = true; 
                    return;
                }
            }
            e.HasMorePages = false;
        }

        private void відобразитиСтатистикуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FStats f = new FStats();
            f.SetDataTable(MySklad.GetStatisticsByGroup());
            f.ShowDialog();
        }

        private void роздрукуватиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }
    }
}
