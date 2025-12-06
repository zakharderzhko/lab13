using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace Sklad
{
    public class TSklad
    {
        public DataTable TabSklad = new DataTable();

        public DataView SkladView;

        public string FiltrCriteria; // оголошуємо поле для зберігання критерію фільтрування
        public string SortCriteria; // оголошуємо поле для зберігання критерію сортування

        public DataGridViewComboBoxColumn cGrupaCB;
        public DataTable DovGrupa = new DataTable();
        public DataTable DovPostachalnyk = new DataTable();
        public DataTable DovOdiVym = new DataTable();
        public DataTable DovSklady = new DataTable();


        public TSklad()
        {
            DataColumn cNpp = new DataColumn("N_пп");
            DataColumn cSklad = new DataColumn("Склад");
            DataColumn cNameGroup = new DataColumn("Група");
            DataColumn cNameProduct = new DataColumn("Назва");
            DataColumn cProduser = new DataColumn("Виробник");
            DataColumn cPost = new DataColumn("Постачальник");
            DataColumn cPrise = new DataColumn("Ціна");
            DataColumn cCount = new DataColumn("Кількість");
            DataColumn cOdi = new DataColumn("Одиниця");
            DataColumn cVartist = new DataColumn("Вартість");


            // Довідник постачальників
            DovPostachalnyk.Columns.Add("Постачальник", typeof(string));
            DovPostachalnyk.Rows.Add("Епіцентр");
            DovPostachalnyk.Rows.Add("Алло");
            DovPostachalnyk.Rows.Add("Фокстрот");
            DovPostachalnyk.Rows.Add("Розетка");

            // Довідник одиниць виміру
            DovOdiVym.Columns.Add("Одиниця", typeof(string));
            DovOdiVym.Rows.Add("шт");
            DovOdiVym.Rows.Add("кг");
            DovOdiVym.Rows.Add("л");

            // Довідник складів
            DovSklady.Columns.Add("Склад", typeof(string));
            DovSklady.Rows.Add("Склад №1");
            DovSklady.Rows.Add("Склад №2");
            DovSklady.Rows.Add("Склад №3");

            cNpp.DataType = System.Type.GetType("System.Int32");
            cSklad.DataType = System.Type.GetType("System.String");
            cNameGroup.DataType = System.Type.GetType("System.String");
            cProduser.DataType = System.Type.GetType("System.String");
            cPost.DataType = System.Type.GetType("System.String");
            cPrise.DataType = System.Type.GetType("System.Decimal");
            cCount.DataType = System.Type.GetType("System.Int32");
            cOdi.DataType = System.Type.GetType("System.String");
            cVartist.DataType = System.Type.GetType("System.Decimal");

            TabSklad.Columns.Add(cNpp);
            TabSklad.Columns.Add(cSklad);
            TabSklad.Columns.Add(cNameGroup);
            TabSklad.Columns.Add(cNameProduct);
            TabSklad.Columns.Add(cProduser);
            TabSklad.Columns.Add(cPost);
            TabSklad.Columns.Add(cPrise);
            TabSklad.Columns.Add(cCount);
            TabSklad.Columns.Add(cOdi);
            TabSklad.Columns.Add(cVartist);

            SkladView = new DataView(TabSklad); // створюємо змінну типу DataView для можливості фільтрування і сортування
        }

        public void TSkladAddRow(string pNameGroup, string pSklad, string pNameProduct, string pProduser, string pPost, int pCount, decimal pPrise, string pOdi)
        {
            int npp = 1;
            if (TabSklad.Rows.Count > 0)
            {
                try
                {
                    object last = TabSklad.Rows[TabSklad.Rows.Count - 1]["N_пп"];
                    if (last != DBNull.Value) npp = Convert.ToInt32(last) + 1;
                }
                catch { npp = TabSklad.Rows.Count + 1; }
            }

            DataRow row = TabSklad.NewRow();
            row["N_пп"] = npp;
            row["Група"] = pNameGroup ?? "";
            row["Склад"] = pSklad ?? "";
            row["Назва"] = pNameProduct ?? "";
            row["Виробник"] = pProduser ?? "";
            row["Постачальник"] = pPost ?? "";
            row["Кількість"] = pCount;
            row["Одиниця"] = pOdi ?? "";
            row["Ціна"] = pPrise;
            row["Вартість"] = pCount * pPrise;

            TabSklad.Rows.Add(row);
        }

        public void ColumnPropSet(DataGridView DGV)
        {
            if (DGV.Columns["N_пп"] != null)
                DGV.Columns["N_пп"].HeaderText = "№ п/п";

            if (DGV.Columns["Група"] != null)
                DGV.Columns["Група"].HeaderText = "Група";

            if (DGV.Columns["Назва"] != null)
                DGV.Columns["Назва"].HeaderText = "Назва";

            if (DGV.Columns["Виробник"] != null)
                DGV.Columns["Виробник"].HeaderText = "Виробник";

            if (DGV.Columns["Кількість"] != null)
            {
                DGV.Columns["Кількість"].HeaderText = "Кількість";
                DGV.Columns["Кількість"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            if (DGV.Columns["Ціна"] != null)
            {
                DGV.Columns["Ціна"].HeaderText = "Ціна";
                DGV.Columns["Ціна"].DefaultCellStyle.Format = "N2";
                DGV.Columns["Ціна"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            if (DGV.Columns["Вартість"] != null)
            {
                DGV.Columns["Вартість"].HeaderText = "Вартість";
                DGV.Columns["Вартість"].DefaultCellStyle.Format = "N2";
                DGV.Columns["Вартість"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        public void ZapTabFile() // ZapTabFile – метод для записування таблиці у текстовий файл із назвою FtabSklad.txt
        {
            string sNameFile, textRow; 
            // Визначаємо, за допомогою методу Directory.GetCurrentDirectory(), ім’я каталогу із *.exe-файлом проекту, для 
            // наступного запису туди нашої таблиці
            string sdir = Directory.GetCurrentDirectory(); sNameFile = sdir + @"\FTabSklad.txt"; // це шлях і ім’я файлу, куди ми будемо записувати таблицю
            try // проконтролюємо наявність подібного файлу
            {
                if (File.Exists(sNameFile)) // якщо такий файл уже є, то знищуємо його, щоб створити ще раз
                {
                    File.Delete(sNameFile);
                } // StreamWriter – клас для введення записів у текстовий файл // Створюємо екземпляр класу StreamWriter, який записуватиме рядки
                using (StreamWriter sw = new StreamWriter(sNameFile))
                {
                    foreach (DataRow rr in TabSklad.Rows)
                    { /* Для кожного рядка rr із таблиці TabSklad побудуємо текстовий образ рядка, об’єднуючи значення полів таблиці у один рядок і розділяючи їх символом “;” */
                        textRow = rr["Група"] + ";" + rr["Склад"] + ";" + rr["Назва"] + ";" + rr["Виробник"] + ";" +
                        rr["Постачальник"] + ";" + Convert.ToString(rr["Кількість"]) + ";" + rr["Одиниця"] + ";" + Convert.ToString(rr["Ціна"]);
                        // записуємо рядок у файл
                        sw.WriteLine(textRow);
                    }
                }
            } 
            catch (Exception e) // якщо не вдалось записати, видаємо повідомлення про помилку
            {
                MessageBox.Show("Таблиця не записана");
            }
        }

        public void ReadTabFile(DataGridView DGS)
        {
            string sNameFile;
            string sdir = Directory.GetCurrentDirectory();
            sNameFile = sdir + @"\FTabSklad.txt";

            TabSklad.Rows.Clear();

            using (StreamReader sr = new StreamReader(sNameFile))
            {
                while (!sr.EndOfStream)
                {
                    string textRow = sr.ReadLine().Trim();
                    if (string.IsNullOrWhiteSpace(textRow)) continue;

                    string[] parts = textRow.Split(';');

                    if (parts.Length != 8)
                        continue;

                    string pGrupa = parts[0].Trim();
                    string pSklad = parts[1].Trim();
                    string pNazva = parts[2].Trim();
                    string pVyrobnyk = parts[3].Trim();
                    string pPostachalnyk = parts[4].Trim();
                    int pKilkist = 0;
                    int.TryParse(parts[5].Trim(), out pKilkist);
                    string pOdiVym = parts[6].Trim();
                    decimal PCina = 0;
                    decimal.TryParse(parts[7].Trim().Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out PCina);

                    // Додаємо рядок через нову сигнатуру
                    TSkladAddRow(pGrupa, pSklad, pNazva, pVyrobnyk, pPostachalnyk, pKilkist, PCina, pOdiVym);
                }
            }

            SetSumy(DGS);


        }


        public void TSkladValFiltr(String PFilter, DataGridView DGV) // TSkladValFiltr – метод для накладання на нашу таблицю і, відповідно, на DataGridView, фільтру по значеннях стовпців
        {
            try
            {
                SkladView.RowFilter = PFilter; // встановлюємо значення фільтру
                FiltrCriteria = PFilter; DGV.DataSource = SkladView; // призначаємо гріду – SkladView, як джерело даних
            }
            catch
            {
                MessageBox.Show("Введений Вами Фільтр не правильний");
                return;
            }
        } // TSkladValSort – метод для встановлення порядку сортування
        public void TSkladValSort(String PSort, DataGridView DGV, DataGridView DGVSum)
        {
            try
            {
                SkladView.Sort = PSort; // встановлюємо критерій сортування
                SortCriteria = PSort; DGV.DataSource = SkladView; // призначаємо гріду – SkladView, як джерело даних
                DGV.Refresh();
            }
            catch
            {
                MessageBox.Show("Введений Вами критерій сортування не правильний");
                return;
            }
        }

        public void SeekNazva(string sNazva, DataGridView DGV)
        {
            int nn; // nn – номер шуканого рядка
            nn = -5; // від'ємне значення у нас означає, що рядок з шуканою назвою не знайдено
            for (int i = 0; i < DGV.Rows.Count; i++) // для кожного рядка rr із DGV
            {
                if ((string)DGV.Rows[i].Cells["Назва"].Value == sNazva)
                {
                    nn = i; // якщо назви збіглися, то - шуканий рядок знайдено. Записуємо його номер у nn і виходимо з циклу
                    break;
                }
            }
            if (nn >= 0) // якщо рядок знайдено, то показуємо його і виділяємо
            {
                DGV.FirstDisplayedCell = DGV.Rows[nn].Cells["Назва"]; DGV.Rows[nn].Selected = true; // виділяємо знайдений рядок DGV.CurrentCell = DGV.Rows[nn].Cells["Назва"]; // встановлюємо знайдений рядок як поточний
            }
            else
            {
                MessageBox.Show("Значення не знайдено");
            }
        }

        public void SetSumy(DataGridView DGV) // Створюємо таблицю для підсумків, записуємо у неї підсумки і призначимо цю таблицю джерелом даних для DGV
        {
            string sGrupa, ssort; 
            int i;
            decimal DSuma;
            DataTable TabSkladSum = new DataTable(); // оголошуємо public-змінну TabSkladSum типу DataTable // Таблиця підсумків буде складатись із 2-х стовпців: Група та Вартість
            DataColumn cNameGroupS = new DataColumn("Група");
            DataColumn cVartistS = new DataColumn("Вартість"); // Оголошуємо типи даних, які будуть зберігатись у стовпцях
            cNameGroupS.DataType = System.Type.GetType("System.String");
            cVartistS.DataType = System.Type.GetType("System.Decimal"); // Додаємо стовпці до таблиці
            TabSkladSum.Columns.Add(cNameGroupS);
            TabSkladSum.Columns.Add(cVartistS); ssort = SkladView.Sort; // запам’ятовуємо заданий користувачем критерій сортування SkladView.Sort = "Група"; // встановлюємо сортування по групах товару. SkladView.Count – кількість рядків
            i = 0; while (i < SkladView.Count) // цикл для всіх рядків із таблиці TabSklad, що впорядкована по групах
            {
                sGrupa = (string)SkladView[i]["Група"]; // обираємо чергову групу товару
                DSuma = 0.0M; // обнулюємо значення суми вартостей для кожної групи
                while ((i < SkladView.Count) & (sGrupa == (string)SkladView[i]["Група"]))
                {
                    try // якщо у якомусь рядку не записана вартість, то скористаємось засобами try - catch
                    {
                        DSuma = DSuma + (decimal)SkladView[i]["Вартість"]; // накопичуємо суму вартостей по групі
                    }
                    catch
                    {
                        SkladView[i]["Вартість"] = 0M;
                    }
                    i = i + 1;
                    if (i == SkladView.Count) { break; }
                }
                DataRow rowSkladSum = TabSkladSum.NewRow(); // створюємо новий рядок у таблиці підсумків
                rowSkladSum["Група"] = sGrupa; // записуємо значення назви групи
                rowSkladSum["Вартість"] = DSuma; // записуємо значення суми вартостей по групі
                TabSkladSum.Rows.Add(rowSkladSum); // додаємо сформований рядок до таблиці підсумків
            }
            DGV.DataSource = TabSkladSum; // призначаємо TabSkladSum джерелом даних для гріда // Відновлюємо критерій сортування, оскільки для сум було встановлено сортування по групі
            SkladView.Sort = SortCriteria;
        }

        // Створюємо таблицю для довідника груп, записуємо у неї дані звичайним присвоєнням
        public void CreateDovGrupa()
        {
            //DataTable DovGrupa = new DataTable();
            // Оголошуємо public-змінну TabSkladSum типу DataTable
            // Таблиця DovGrupa буде складатись із 1 стовпця – назва групи товару.
            DataColumn cNameGroup = new DataColumn("Група");
            // Оголошуємо тип даних, які будуть зберігатись у стовпці
            cNameGroup.DataType = System.Type.GetType("System.String");
            DovGrupa.Columns.Add(cNameGroup); // додаємо стовпець до таблиці
            DataRow rowSklad0 = DovGrupa.NewRow();
            rowSklad0[cNameGroup] = "Книги";
            DovGrupa.Rows.Add(rowSklad0); // додаємо сформований рядок до таблиці
            DataRow rowSklad1 = DovGrupa.NewRow();
            rowSklad1[cNameGroup] = "CD";
            DovGrupa.Rows.Add(rowSklad1); // додаємо сформований рядок до таблиці
            DataRow rowSklad2 = DovGrupa.NewRow();
            rowSklad2[cNameGroup] = "DVD";
            DovGrupa.Rows.Add(rowSklad2); // додаємо сформований рядок до таблиці
            DataRow rowSklad3 = DovGrupa.NewRow();
            rowSklad3[cNameGroup] = "Мобілки";
            DovGrupa.Rows.Add(rowSklad3); // додаємо сформований рядок до таблиці
            DataRow rowSklad4 = DovGrupa.NewRow();
            rowSklad4[cNameGroup] = "Плеєри";
            DovGrupa.Rows.Add(rowSklad4); // додаємо сформований рядок до таблиці
            DataRow rowSklad5 = DovGrupa.NewRow();
            rowSklad5[cNameGroup] = "Аксессуари";
            DovGrupa.Rows.Add(rowSklad5); // додаємо сформований рядок до таблиці
            DataRow rowSklad6 = DovGrupa.NewRow();
            rowSklad6[cNameGroup] = "Дисплеї";
            DovGrupa.Rows.Add(rowSklad6); // додаємо сформований рядок до таблиці
            DataRow rowSklad7 = DovGrupa.NewRow();
            rowSklad7[cNameGroup] = "Корпуси";
            DovGrupa.Rows.Add(rowSklad7); // додаємо сформований рядок до таблиці
            DataRow rowSklad8 = DovGrupa.NewRow();
            rowSklad8[cNameGroup] = "Блоки живлення";
            DovGrupa.Rows.Add(rowSklad8); // додаємо сформований рядок до таблиці
            DataRow rowSklad9 = DovGrupa.NewRow();
            rowSklad9[cNameGroup] = "Клавіатури";
            DovGrupa.Rows.Add(rowSklad9); // додаємо сформований рядок до таблиці
            int nn = DovGrupa.Rows.Count;
        }

        public void AddComboGrupa(DataGridView DGV)
        {
            /* Цей метод замінить стовбець назв груп товарів у гріді таблиці складу на новий стовбець. 
            Від старого він буде відрізнятися тим, що при клацанні на його комірку буде відкриватись 
            вікно-випадайка, з якого користувач зможе вибрати групу товару і занести її у комірку 
            простим клацанням мишки. Окрім спрощення процесу введення даних, такий спосіб іще гарантує, 
            що у комірку завжди буде введено лише допустиме значення. */
            // Створимо екземпляр DataGridViewComboBoxColumn
            DataGridViewComboBoxColumn cGrupaCB = new DataGridViewComboBoxColumn();
            // Джерело даних для нового стовпця – стовпець з іменем "Група" у таблиці – джерелі даних.
            cGrupaCB.DataPropertyName = "Група";
            cGrupaCB.Name = "cNameGroupComb"; // Назва нового стовпця
            cGrupaCB.HeaderText = "Група"; // Заголовок на гріді нового стовпця
            cGrupaCB.DropDownWidth = 200; // Ширина "випадайки"
            cGrupaCB.Width = 120; // Ширина стовпця
            cGrupaCB.MaxDropDownItems = 7; // Кількість рядків випадайки, які одночасно будуть видимі
            cGrupaCB.FlatStyle = FlatStyle.Flat;
            cGrupaCB.ValueType = System.Type.GetType("System.string"); // Тип даних нового стовпця
            String s; Int32 n;
            n = DovGrupa.Rows.Count;
            // Тут DovGrupa – таблиця-довідник, яка містить допустимі значення назви групи товару

            cGrupaCB.Items.Clear();

            foreach (DataRow r in DovGrupa.Rows) // Для кожного рядка rr із таблиці DovGrupa
            {
                s = (string)r["Група"];
                cGrupaCB.Items.Add(r["Група"]); // Додаємо у "випадайку" елементи із довідника
            }
            DGV.Columns.Add(cGrupaCB); // Додаємо новий стовбець до гріда
                                       // Перезаписати значення комірок із старого стовпця у новий
            String ss;
            // Для кожного рядка rr контрола DGV типу DataGridView
            foreach (DataGridViewRow rrr in DGV.Rows)
            {
                ss = (string)rrr.Cells["Група"].Value; // ss – значення комірки назви групи у старому стовпці
                                                       // Перезаписуєм значення комірки старого стовпця у комірку нового стовпця
                rrr.Cells["Група"].Value = rrr.Cells["Група"].Value;
            }
            DGV.Columns.Remove("Група"); // Вилучаємо старий стовбець типу DataGridViewTextBoxColumn
            DGV.Columns["cNameGroupComb"].Name = "Група"; // Перейменовуємо новий у старе ім’я
            DGV.Columns["Група"].DisplayIndex = 1; // Встановлюємо порядок виводу стовпця у гріді
            ((DataGridViewComboBoxColumn)DGV.Columns["Група"]).DataSource = null;
            ((DataGridViewComboBoxColumn)DGV.Columns["Група"]).DisplayMember = "";
            ((DataGridViewComboBoxColumn)DGV.Columns["Група"]).ValueMember = "";
            DGV.Refresh(); // Оновлюємо DataGridView
        }

        public void AddComboPostachalnyk(DataGridView DGV)
        {
            DataGridViewComboBoxColumn cPostachalnykCB = new DataGridViewComboBoxColumn();

            cPostachalnykCB.DataPropertyName = "Постачальник";
            cPostachalnykCB.Name = "cPostachalnykComb";
            cPostachalnykCB.HeaderText = "Постачальник";
            cPostachalnykCB.DropDownWidth = 250;
            cPostachalnykCB.Width = 140;
            cPostachalnykCB.MaxDropDownItems = 8;
            cPostachalnykCB.FlatStyle = FlatStyle.Flat;
            cPostachalnykCB.ValueType = typeof(string);

            cPostachalnykCB.Items.Clear();

            foreach (DataRow r in DovPostachalnyk.Rows)
            {
                cPostachalnykCB.Items.Add(r["Постачальник"]);
            }

            DGV.Columns.Add(cPostachalnykCB);

            foreach (DataGridViewRow row in DGV.Rows)
            {
                if (row.Cells["Постачальник"].Value != null)
                    row.Cells["Постачальник"].Value = row.Cells["Постачальник"].Value;
            }

            DGV.Columns.Remove("Постачальник");
            DGV.Columns["cPostachalnykComb"].DisplayIndex = 4;  
            DGV.Columns["cPostachalnykComb"].Name = "Постачальник";
            ((DataGridViewComboBoxColumn)DGV.Columns["Постачальник"]).DataSource = null;
            ((DataGridViewComboBoxColumn)DGV.Columns["Постачальник"]).DisplayMember = "";
            ((DataGridViewComboBoxColumn)DGV.Columns["Постачальник"]).ValueMember = "";
            DGV.Refresh();
        }

        public void AddComboOdiVym(DataGridView DGV)
        {
            DataGridViewComboBoxColumn cOdiVymCB = new DataGridViewComboBoxColumn();

            cOdiVymCB.DataPropertyName = "Одиниця";
            cOdiVymCB.Name = "cOdiVymComb";
            cOdiVymCB.HeaderText = "Одиниця";
            cOdiVymCB.DropDownWidth = 120;
            cOdiVymCB.Width = 80;
            cOdiVymCB.MaxDropDownItems = 6;
            cOdiVymCB.FlatStyle = FlatStyle.Flat;
            cOdiVymCB.ValueType = typeof(string);

            cOdiVymCB.Items.Clear();

            foreach (DataRow r in DovOdiVym.Rows)
            {
                cOdiVymCB.Items.Add(r["Одиниця"]);
            }

            DGV.Columns.Add(cOdiVymCB);

            foreach (DataGridViewRow row in DGV.Rows)
            {
                if (row.Cells["Одиниця"].Value != null)
                    row.Cells["Одиниця"].Value = row.Cells["Одиниця"].Value;
            }

            DGV.Columns.Remove("Одиниця");
            DGV.Columns["cOdiVymComb"].DisplayIndex = 6;
            DGV.Columns["cOdiVymComb"].Name = "Одиниця";
            DGV.Refresh();
        }

        public void UpdateAllCombos(DataGridView DGV)
        {
            var colGr = DGV.Columns["Група"] as DataGridViewComboBoxColumn;
            if (colGr != null)
            {
                colGr.Items.Clear();
                foreach (DataRow r in DovGrupa.Rows)
                    colGr.Items.Add(r["Група"]);
            }

            var colPost = DGV.Columns["Постачальник"] as DataGridViewComboBoxColumn;
            if (colPost != null)
            {
                colPost.Items.Clear();
                foreach (DataRow r in DovPostachalnyk.Rows)
                    colPost.Items.Add(r["Постачальник"]);
            }

            var colOdi = DGV.Columns["Одиниця"] as DataGridViewComboBoxColumn;
            if (colOdi != null)
            {
                colOdi.Items.Clear();
                foreach (DataRow r in DovOdiVym.Rows)
                    colOdi.Items.Add(r["Одиниця"]);
            }

            var colSkl = DGV.Columns["Склад"] as DataGridViewComboBoxColumn;
            if (colSkl != null)
            {
                colSkl.Items.Clear();
                foreach (DataRow r in DovSklady.Rows)
                    colGr.Items.Add(r["Склад"]);
            }

            DGV.Refresh();
        }

        public DataTable GetStatisticsByGroup()
        {
            var dt = new DataTable();
            dt.Columns.Add("Група", typeof(string));
            dt.Columns.Add("Кількість_позицій", typeof(int));
            dt.Columns.Add("Загальна_вартість", typeof(decimal));
            dt.Columns.Add("Середня_ціна", typeof(decimal));
            dt.Columns.Add("Мін_ціна", typeof(decimal));
            dt.Columns.Add("Макс_ціна", typeof(decimal));

            var q = TabSklad.AsEnumerable()
                    .GroupBy(r => r.Field<string>("Група"))
                    .Select(g => new
                    {
                        Group = g.Key,
                        Count = g.Count(),
                        Sum = g.Sum(r => r.Field<decimal>("Вартість")),
                        Avg = g.Average(r => r.Field<decimal>("Ціна")),
                        Min = g.Min(r => r.Field<decimal>("Ціна")),
                        Max = g.Max(r => r.Field<decimal>("Ціна"))
                    });

            foreach (var item in q)
            {
                var row = dt.NewRow();
                row["Група"] = item.Group;
                row["Кількість_позицій"] = item.Count;
                row["Загальна_вартість"] = item.Sum;
                row["Середня_ціна"] = item.Avg;
                row["Мін_ціна"] = item.Min;
                row["Макс_ціна"] = item.Max;
                dt.Rows.Add(row);
            }

            return dt;
        }

    }
}
