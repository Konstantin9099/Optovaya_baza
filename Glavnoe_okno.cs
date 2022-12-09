using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Word = Microsoft.Office.Interop.Word;

namespace Baza
{
    public partial class Glavnoe_okno : Form
    {
        public int ID = 0;
        public string itog;

        public Glavnoe_okno(int ID_log)
        {
            InitializeComponent();
            GetInfo(ID_log);
            ID = ID_log;
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            this.textBox1.ForeColor = System.Drawing.Color.Navy;
            this.textBox3.ForeColor = System.Drawing.Color.Navy;
            this.textBox3.MaxLength = 12;
            this.textBox4.ForeColor = System.Drawing.Color.Navy;
            this.textBox5.ForeColor = System.Drawing.Color.Navy;
            this.textBox5.MaxLength = 12;
            textBox5.KeyPress += (sender, e) => e.Handled = true;
            this.textBox6.ForeColor = System.Drawing.Color.Navy;
            textBox6.KeyPress += (sender, e) => e.Handled = true;
            this.textBox7.ForeColor = System.Drawing.Color.Navy;
            this.textBox8.ForeColor = System.Drawing.Color.Navy;
            this.textBox9.ForeColor = System.Drawing.Color.Navy;
            textBox9.KeyPress += (sender, e) => e.Handled = true;
            this.textBox10.ForeColor = System.Drawing.Color.Navy;
            textBox10.KeyPress += (sender, e) => e.Handled = true;
            this.textBox11.ForeColor = System.Drawing.Color.Navy;
            this.textBox12.ForeColor = System.Drawing.Color.Navy;
            textBox11.KeyPress += (sender, e) => e.Handled = true;
            textBox13.KeyPress += (sender, e) => e.Handled = true;
            this.textBox13.ForeColor = System.Drawing.Color.Navy;
            this.textBox16.ForeColor = System.Drawing.Color.Navy;
            this.textBox17.ForeColor = System.Drawing.Color.Navy;
            textBox18.KeyPress += (sender, e) => e.Handled = true;
            this.textBox18.ForeColor = System.Drawing.Color.Navy;
            this.comboBox1.ForeColor = System.Drawing.Color.Navy;
            comboBox1.KeyPress += (sender, e) => e.Handled = true;
            this.comboBox2.ForeColor = System.Drawing.Color.Navy;
            comboBox2.KeyPress += (sender, e) => e.Handled = true;
            this.comboBox3.ForeColor = System.Drawing.Color.Navy;
            comboBox4.KeyPress += (sender, e) => e.Handled = true;
            this.comboBox5.ForeColor = System.Drawing.Color.Navy;
            this.comboBox4.ForeColor = System.Drawing.Color.Navy;
            this.comboBox6.ForeColor = System.Drawing.Color.Navy;
            comboBox8.KeyPress += (sender, e) => e.Handled = true;
        }

        private void Glavnoe_okno_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //Функция, позволяющая отправить команду на сервер БД для оптимизации кода.
        public void Action(string query)
        {
            MySqlConnection conn = DBUtils.GetDBConnection();
            MySqlCommand cmDB = new MySqlCommand(query, conn);
            try
            {
                conn.Open();
                cmDB.ExecuteReader();
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
            }
        }

        // Получаем из БД данные для таблиц программы и выводим их в DataGridView.
        public void GetInfo(int ID)
        {
            // Вкладка "Клиенты" - Таблица "Данные клиентов".
            string query = "SELECT id_klienta AS 'Код клиента', naimenovanie_klienta AS 'Наименование клиента', inn_klienta AS 'ИНН', telefon AS 'Номер телефона' FROM Klient;";
            MySqlConnection conn = DBUtils.GetDBConnection();
            MySqlDataAdapter sda = new MySqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                dataGridView1.DataSource = dt;
                dataGridView1.ClearSelection();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.ClearSelection();
                this.dataGridView1.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView1.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[1].Width = 300;
                this.dataGridView1.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[2].Width = 155;
                this.dataGridView1.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[3].Width = 155;
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }

            // Вкладка "Заказы" - Таблица "Заказы клиентов".
            string query1 = "select zakaz.id_zakaz as 'Номер заказа', Zakaz.data_zakaza as 'Дата заказа', tovar_id as 'Код товара', sklad_id as 'Склад', naimenovanie_tovara as 'Наименование товара', ed_izmer_tov as 'Ед. измер.', stoim_ed_tov as 'Цена', Zakaz.kolichestvo_tovara as 'Кол-во', Zakaz.stoimost_zakaza as 'Сумма', Klient.naimenovanie_klienta 'Клиент', Klient.inn_klienta as 'ИНН', Klient.telefon as 'Номер телефона' from tovar_sklad, tovar, sklad, zakaz, klient where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and Zakaz.tovar_sklad_id=Tovar_sklad.id_tovar_sklad and Zakaz.klient_id=Klient.id_klienta;";
            MySqlConnection conn1 = DBUtils.GetDBConnection();
            MySqlDataAdapter sda1 = new MySqlDataAdapter(query1, conn1);
            DataTable dt1 = new DataTable();
            try
            {
                conn1.Open();
                dataGridView2.DataSource = dt1;
                dataGridView2.ClearSelection();
                sda1.Fill(dt1);
                dataGridView2.DataSource = dt1;
                dataGridView2.ClearSelection();
                this.dataGridView2.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView2.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[0].Width = 70;
                this.dataGridView2.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[1].Width = 90;
                this.dataGridView2.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[2].Width = 70;
                this.dataGridView2.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[3].Width = 70;
                this.dataGridView2.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[4].Width = 270;
                this.dataGridView2.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[5].Width = 50;
                this.dataGridView2.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[6].Width = 100;
                this.dataGridView2.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[7].Width = 100;
                this.dataGridView2.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView2.Columns[8].Width = 100;
                dataGridView2.Columns[9].Visible = false;
                dataGridView2.Columns[10].Visible = false;
                dataGridView2.Columns[11].Visible = false;
                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn1.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }

            // Вкладка "Оплаты" - Таблица "Оплаты клиентов".
            string query2 = "SELECT Oplata.id_dokumenta as 'Номер документа', Oplata.data_oplaty as 'Дата оплаты',  Oplata.summa_oplaty as 'Сумма оплаты', Klient.naimenovanie_klienta as 'Наименование клиента', Klient.inn_klienta as 'ИНН клиента', Klient.Telefon as 'Номер телефона' FROM Oplata, Klient where Oplata.klient_id=Klient.id_klienta;";
            MySqlConnection conn2 = DBUtils.GetDBConnection();
            MySqlDataAdapter sda2 = new MySqlDataAdapter(query2, conn2);
            DataTable dt2 = new DataTable();
            try
            {
                conn2.Open();
                dataGridView3.DataSource = dt2;
                dataGridView3.ClearSelection();
                sda2.Fill(dt2);
                dataGridView3.DataSource = dt2;
                dataGridView3.ClearSelection();
                this.dataGridView3.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView3.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[0].Width = 80;
                this.dataGridView3.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[1].Width = 90;
                this.dataGridView3.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[2].Width = 100;
                this.dataGridView3.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[3].Width = 270;
                this.dataGridView3.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[4].Width = 120;
                this.dataGridView3.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView3.Columns[5].Width = 120;
                dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn2.Close();

                decimal Total = 0;

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    Total += Convert.ToDecimal(dataGridView3.Rows[i].Cells[2].Value);
                }

                itog = "Итого поступило: " + Total.ToString("f2") + "руб.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }

            // Вкладка "Товарная номенклатура" - Таблица "Товарная номенклатура".
            //string query3 = "select tovar.id_tovara as 'Код товара', sklad.naimenovanie_sklada as 'Склад', tovar.naimenovanie_tovara as 'Наименование товара', tovar.ed_izmer_tov as 'Ед. измер.', tovar.stoim_ed_tov as 'Цена' from tovar, sklad, tovar_sklad where tovar.id_tovara=tovar_sklad.tovar_id and sklad.id_sklada=tovar_sklad.sklad_id;";
            string query3 = "select tovar.id_tovara as 'Код товара', tovar.naimenovanie_tovara as 'Наименование товара', tovar.ed_izmer_tov as 'Ед. измер.', tovar.stoim_ed_tov as 'Цена' from tovar order by tovar.naimenovanie_tovara;";
            MySqlConnection conn3 = DBUtils.GetDBConnection();
            MySqlDataAdapter sda3 = new MySqlDataAdapter(query3, conn3);
            DataTable dt3 = new DataTable();
            try
            {
                conn3.Open();
                dataGridView7.DataSource = dt3;
                dataGridView7.ClearSelection();
                sda3.Fill(dt3);
                dataGridView7.DataSource = dt3;
                dataGridView7.ClearSelection();
                this.dataGridView7.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView7.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView7.Columns[0].Width = 70;
                this.dataGridView7.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView7.Columns[1].Width = 310;
                this.dataGridView7.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView7.Columns[2].Width = 50;
                this.dataGridView7.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView7.Columns[3].Width = 100;
                dataGridView7.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn3.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }

            // Вкладка "Тованая номенклатура" - Таблица "Склады".
            string query4 = "select sklad.id_sklada as 'Код склада', sklad.naimenovanie_sklada as 'Наименование склада', sklad.fio_zavsklada as 'Ф.И.О. завсклада', sklad.telefon as 'Телефон' from sklad;";
            MySqlConnection conn4 = DBUtils.GetDBConnection();
            MySqlDataAdapter sda4 = new MySqlDataAdapter(query4, conn4);
            DataTable dt4 = new DataTable();
            try
            {
                conn4.Open();
                dataGridView8.DataSource = dt4;
                dataGridView8.ClearSelection();
                sda4.Fill(dt4);
                dataGridView8.DataSource = dt4;
                dataGridView8.ClearSelection();
                this.dataGridView8.ForeColor = System.Drawing.Color.Navy;
                //dataGridView8.Columns[0].Visible = false;
                this.dataGridView8.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView8.Columns[0].Width = 70;
                this.dataGridView8.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView8.Columns[1].Width = 230;
                this.dataGridView8.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView8.Columns[2].Width = 250;
                this.dataGridView8.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView8.Columns[3].Width = 110;
                dataGridView8.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn4.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }

            // Вкладка "Тованая номенклатура" - Таблица "Перемещение товаров".
            string query5 = "select tovar.id_tovara as 'Код товара', tovar.naimenovanie_tovara as 'Наименование товара', sklad.naimenovanie_sklada as 'Склад' from sklad, tovar, tovar_sklad where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara order by tovar.naimenovanie_tovara;";
            //string query5 = "select tovar.id_tovara as 'Код товара', tovar.naimenovanie_tovara as 'Наименование товара' from tovar order by tovar.naimenovanie_tovara;";
            MySqlConnection conn5 = DBUtils.GetDBConnection();
            MySqlDataAdapter sda5 = new MySqlDataAdapter(query5, conn5);
            DataTable dt5 = new DataTable();
            try
            {
                conn5.Open();
                dataGridView9.DataSource = dt5;
                dataGridView9.ClearSelection();
                sda5.Fill(dt5);
                dataGridView9.DataSource = dt5;
                dataGridView9.ClearSelection();
                this.dataGridView9.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView9.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView9.Columns[0].Width = 70;
                this.dataGridView9.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView9.Columns[1].Width = 250;
                this.dataGridView9.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView9.Columns[2].Width = 200;
                dataGridView9.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn5.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }
        }

        // Вывод выпадающих списков в элементы управления ComboBox.
        private void Glavnoe_okno_Load(object sender, EventArgs e)
        {
            // Список клиентов во вкладке "Клиенты".
            try
            {
                string query = "select Klient.naimenovanie_klienta from Klient order by Klient.naimenovanie_klienta;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.DropDownHeight = 150;
                    comboBox1.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Список товаров во вкладке "Заказы".
            try
            {
                string query = "select Tovar.naimenovanie_tovara from Tovar order by naimenovanie_tovara;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox2.DropDownHeight = 150;
                    comboBox2.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Список клиентов во вкладке "Оплаты".
            try
            {
                string query = "select Klient.naimenovanie_klienta from Klient order by Klient.naimenovanie_klienta;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox4.DropDownHeight = 150;
                    comboBox4.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Список "Наименование товаров" во вкладке "Склады".
            try
            {
                string query = "select tovar.naimenovanie_tovara from tovar order by tovar.naimenovanie_tovara;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox3.DropDownHeight = 150;
                    comboBox3.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Список "Склады" во вкладке "Склады".
            try
            {
                string query = "select sklad.naimenovanie_sklada from sklad order by sklad.naimenovanie_sklada;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox6.DropDownHeight = 150;
                    comboBox6.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Список складов во вкладке "Товарная номенклатура".
            try
            {
                string query = "select sklad.naimenovanie_sklada from sklad;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(query, conn);
                MySqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox8.DropDownHeight = 150;
                    comboBox8.Items.Add(reader.GetString(0));
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // ******************* ВКЛАДКА "КЛИЕНТЫ" *********************
        //                 Таблица "ДАННЫЕ КЛИЕНТОВ"
        // Вывод данных в текстовые поля из dataGridView1.
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            this.textBox2.ForeColor = System.Drawing.Color.Navy;
            textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            this.textBox3.ForeColor = System.Drawing.Color.Navy;
            maskedTextBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
        }

        // Запрет на ввод в поле ввода номер телефона любых букв и символов, кроме цифр и клавиши backspace.
        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
                e.Handled = true;
        }

        // Запрет на ввод в поле ввода ИНН любых букв и символов, кроме цифр и клавиши backspace.
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
                e.Handled = true;
        }

        // Кнопка "Добавить данные клиента".
        private void button1_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (textBox2.Text == null || textBox3.Text == "")
                MessageBox.Show(
                    "Введите данные.",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {
                string klient = textBox2.Text;
                DialogResult res = MessageBox.Show($"Вы уверены, что хотите добавить клиента:\n\n{klient}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = "INSERT INTO Klient (naimenovanie_klienta, inn_klienta, telefon) VALUES ('" + textBox2.Text + "', '" + textBox3.Text + "', '" + maskedTextBox1.Text + "'); ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    textBox2.Clear();
                    textBox3.Clear();
                    maskedTextBox1.Clear();
                }
            }
        }

        // Кнопка "Изменить данные клиента".
        private void button2_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода/вывода данных.
            if (textBox2.Text == null || textBox3.Text == "")
                MessageBox.Show(
                    "Не заполнены поля ввода.",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {
                string name = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                DialogResult res = MessageBox.Show($"Вы уверены, что хотите изменить данные клиента:\n\n{name}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    int id_klient = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                    string query = " UPDATE Klient SET naimenovanie_klienta='" + textBox2.Text + "', inn_klienta='" + textBox3.Text + "', telefon='" + maskedTextBox1.Text + "' WHERE id_klienta=" + id_klient + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    textBox2.Clear();
                    textBox3.Clear();
                    maskedTextBox1.Clear();
                }
            }
        }

        // Кнопка "Найти клиента".
        private void button3_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы была заполнена строка поиска.
            if (textBox1.Text == null || textBox1.Text == "")
                MessageBox.Show(
                    "Вы не ввели запрос в строке поиска!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                            if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()))
                            {
                                dataGridView1.Rows[i].Selected = true;
                                break;
                            }
                }
            }
            textBox1.Clear();
        }

        // Кнопка "Печать данных клиентов".
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                title.Range.Font.Size = 14;
                title.Range.Font.Color = Word.WdColor.wdColorDarkRed;
                title.Range.Text = $"Оптовая база\nДанные клиентов (код, наименование организации, ИНН, номер телефона)";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView1.RowCount, dataGridView1.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 12;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView1.RowCount; i++)
                    for (int j = 1; j <= dataGridView1.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView1.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // ******************* ВКЛАДКА "ЗАКАЗЫ" *********************
        //                 Таблица "ЗАКАЗЫ КЛИЕНТОВ"
        // Вывод данных в текстовые поля из dataGridView1 (Заказы клиентов).
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBox1.Text = dataGridView2.CurrentRow.Cells[9].Value.ToString();
            comboBox2.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
            maskedTextBox2.Text = dataGridView2.CurrentRow.Cells[11].Value.ToString();
            dateTimePicker1.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            textBox5.Text = dataGridView2.CurrentRow.Cells[10].Value.ToString();
            textBox6.Text = dataGridView2.CurrentRow.Cells[6].Value.ToString();
            textBox7.Text = dataGridView2.CurrentRow.Cells[7].Value.ToString();
            textBox8.Text = dataGridView2.CurrentRow.Cells[8].Value.ToString();
            textBox9.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
            textBox10.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            textBox11.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
        }

        // Вывод данных в текстовые поля из dataGridView4 (Выборка заказов).
        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBox1.Text = dataGridView4.CurrentRow.Cells[9].Value.ToString();
            comboBox2.Text = dataGridView4.CurrentRow.Cells[4].Value.ToString();
            maskedTextBox2.Text = dataGridView4.CurrentRow.Cells[11].Value.ToString();
            dateTimePicker1.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();
            textBox5.Text = dataGridView4.CurrentRow.Cells[10].Value.ToString();
            textBox6.Text = dataGridView4.CurrentRow.Cells[6].Value.ToString();
            textBox7.Text = dataGridView4.CurrentRow.Cells[7].Value.ToString();
            textBox8.Text = dataGridView4.CurrentRow.Cells[8].Value.ToString();
            textBox9.Text = dataGridView4.CurrentRow.Cells[2].Value.ToString();
            textBox10.Text = dataGridView4.CurrentRow.Cells[3].Value.ToString();
            textBox11.Text = dataGridView4.CurrentRow.Cells[0].Value.ToString();
        }

        // Автоматическое заполнение полей "ИНН" и "Номер телефона" при выборе клиента в выпадающем списке.
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string klient = comboBox1.Text;
            // ИНН.
            try
            {
                string INN = "select inn_klienta from Klient where naimenovanie_klienta ='" + klient + "';";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(INN, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(INN, conn);
                textBox5.Text = command.ExecuteScalar().ToString();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Номер телефона.
            try
            {
                string telefon = "select telefon from Klient where naimenovanie_klienta ='" + klient + "';";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(telefon, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(telefon, conn);
                maskedTextBox2.Text = command.ExecuteScalar().ToString();
                conn.Close();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Автоматическое заполнение полей "Код товара", "Склад" и "Цена товара" при выборе товара в выпадающем списке.
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string naz_tov = comboBox2.Text;
            // Код товара.
            try
            {
                string kod_tov = "select tovar_id from tovar_sklad, tovar, sklad where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and naimenovanie_tovara = '" + naz_tov + "'; ;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(kod_tov, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(kod_tov, conn);
                textBox9.Text = command.ExecuteScalar().ToString();
                conn.Close();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Склад.
            try
            {
                string sklad = "select sklad_id from tovar_sklad, tovar, sklad where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and naimenovanie_tovara = '" + naz_tov + "'; ;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(sklad, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(sklad, conn);
                textBox10.Text = command.ExecuteScalar().ToString();
                conn.Close();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Цена товара.
            try
            {
                string cena = "select stoim_ed_tov from tovar_sklad, tovar, sklad where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and naimenovanie_tovara = '" + naz_tov + "'; ;";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(cena, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(cena, conn);
                textBox6.Text = command.ExecuteScalar().ToString();
                conn.Close();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }

            textBox7.Clear();
            textBox8.Clear();
        }

        // Кнопка "Печать списка заказов".
        private void button6_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString();
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                title.Range.Font.Size = 12;
                title.Range.Text = $"Оптовая база\nСписок заказов (номер заказа, дата заказа, код товара, склад, наименование, ед. изм., цена, кол-во, сумма, заказчик, ИНН, телефон)" +
                    $"\nДата формирования списка: {date}";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView2.RowCount, dataGridView2.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 7;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView2.RowCount; i++)
                    for (int j = 1; j <= dataGridView2.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView2.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // Кнопка "Найти заказ".
        private void button5_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы была заполнена строка поиска.
            if (textBox4.Text == null || textBox4.Text == "")
                MessageBox.Show(
                    "Вы не ввели запрос в строке поиска!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {

                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    dataGridView2.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                        if (dataGridView2.Rows[i].Cells[j].Value != null)
                            if (dataGridView2.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox4.Text.ToLower()))
                            {
                                dataGridView2.Rows[i].Selected = true;
                                break;
                            }

                }
            }
            textBox4.Clear();
        }

        // Кнопка "Добавить заказ".
        private void button8_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox1.Text.Equals("") || textBox5.Text == null || textBox5.Text == "" || comboBox2.Text.Equals("") || textBox6.Text == null || textBox6.Text == "" || textBox7.Text == null || textBox7.Text == "" || textBox9.Text == null || textBox9.Text == "" || textBox10.Text == null || textBox10.Text == "")
                MessageBox.Show(
                    "Не достаточно данных для оформления заказа!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {
                string klient = comboBox1.Text;
                //string klient = textBox2.Text;
                DialogResult res = MessageBox.Show($"Оформить заказ для клиента: \n\n{klient}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string data = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string naz_klienta = comboBox1.Text;
                    int kod_tov = Int32.Parse(textBox9.Text);
                    int kod_skl = Int32.Parse(textBox10.Text);
                    string query = "INSERT INTO zakaz SET klient_id=(select klient.id_klienta from klient where klient.naimenovanie_klienta='" + naz_klienta + "'), data_zakaza='" + data + "', tovar_sklad_id=(select tovar_sklad.id_tovar_sklad from tovar_sklad where tovar_sklad.sklad_id='" + kod_skl + "' and tovar_sklad.tovar_id='" + kod_tov + "'), kolichestvo_tovara='" + textBox7.Text + "', stoimost_zakaza = ('" + textBox7.Text + "' * (select tovar.stoim_ed_tov from tovar where id_tovara = (select Tovar_sklad.tovar_id from Tovar_sklad, Tovar where Tovar_sklad.tovar_id = Tovar.id_tovara and Tovar_sklad.id_tovar_sklad=(select tovar_sklad.id_tovar_sklad from tovar_sklad where tovar_sklad.sklad_id='" + kod_skl + "' and tovar_sklad.tovar_id='" + kod_tov + "')))); ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    double cena, kol_vo, summ;
                    cena = double.Parse(textBox6.Text);
                    kol_vo = double.Parse(textBox7.Text);
                    summ = cena * kol_vo;
                    textBox8.Text = summ.ToString("f2");
                    textBox11.Clear();
                }
            }
        }

        // Кнопка "Изменить заказ".
        private void button7_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox1.Text.Equals("") || textBox5.Text == null || textBox5.Text == "" || comboBox2.Text.Equals("") || textBox6.Text == null || textBox6.Text == "" || textBox7.Text == null || textBox7.Text == "" || textBox9.Text == null || textBox9.Text == "" || textBox10.Text == null || textBox10.Text == "")
                MessageBox.Show(
                    "Не достаточно данных для оформления заказа!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {
                // Определяем id заказа.
                int nom_zak = Int32.Parse(textBox11.Text);
                string date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                int kod_tov = Int32.Parse(textBox9.Text);
                int kod_skl = Int32.Parse(textBox10.Text);
                DialogResult res = MessageBox.Show($"Изменить данные заказа № {nom_zak}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = " UPDATE zakaz SET klient_id=(select klient.id_klienta from klient where naimenovanie_klienta='" + comboBox1.Text + "'), data_zakaza='" + date + "', tovar_sklad_id=(select tovar_sklad.id_tovar_sklad from tovar_sklad where tovar_sklad.sklad_id='" + kod_skl + "' and tovar_sklad.tovar_id='" + kod_tov + "'), kolichestvo_tovara='" + textBox7.Text + "', stoimost_zakaza = ('" + textBox7.Text + "' * (select tovar.stoim_ed_tov from tovar where id_tovara = (select Tovar_sklad.tovar_id from Tovar_sklad, Tovar where Tovar_sklad.tovar_id = Tovar.id_tovara and Tovar_sklad.id_tovar_sklad=(select tovar_sklad.id_tovar_sklad from tovar_sklad where tovar_sklad.sklad_id='" + kod_skl + "' and tovar_sklad.tovar_id='" + kod_tov + "')))) WHERE id_zakaz=" + nom_zak + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    textBox11.Clear();
                }
            }
        }

        // Выборка заказов по дате.
        private void button10_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string query = "select zakaz.id_zakaz as 'Номер заказа', Zakaz.data_zakaza as 'Дата заказа', tovar_id as 'Код товара', sklad_id as 'Склад', naimenovanie_tovara as 'Наименование товара', ed_izmer_tov as 'Ед. измер.', stoim_ed_tov as 'Цена', Zakaz.kolichestvo_tovara as 'Кол-во', Zakaz.stoimost_zakaza as 'Сумма', Klient.naimenovanie_klienta 'Клиент', Klient.inn_klienta as 'ИНН', Klient.telefon as 'Номер телефона' from tovar_sklad, tovar, sklad, zakaz, klient where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and Zakaz.tovar_sklad_id=Tovar_sklad.id_tovar_sklad and Zakaz.klient_id=Klient.id_klienta and data_zakaza='" + date + "';";
            MySqlConnection conn = DBUtils.GetDBConnection();
            MySqlDataAdapter sda = new MySqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                dataGridView4.DataSource = dt;
                dataGridView4.ClearSelection();
                sda.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.ClearSelection();
                this.dataGridView4.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView4.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[0].Width = 70;
                this.dataGridView4.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[1].Width = 90;
                this.dataGridView4.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[2].Width = 70;
                this.dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[3].Width = 70;
                this.dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[4].Width = 270;
                this.dataGridView4.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[5].Width = 50;
                this.dataGridView4.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[6].Width = 100;
                this.dataGridView4.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[7].Width = 100;
                this.dataGridView4.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[8].Width = 100;
                this.dataGridView4.Columns[9].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[9].Width = 250;
                this.dataGridView4.Columns[10].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[10].Width = 110;
                this.dataGridView4.Columns[11].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[11].Width = 110;
                dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn.Close();
                decimal Total = 0;

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    Total += Convert.ToDecimal(dataGridView4.Rows[i].Cells[8].Value);
                }

                label20.Text = "ИТОГО: " + Total.ToString("f2");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }
        }

        // Выборка заказов по данным клиента.
        private void button9_Click(object sender, EventArgs e)
        {
            string inn_klient = textBox5.Text;
            string query = "select zakaz.id_zakaz as 'Номер заказа', Zakaz.data_zakaza as 'Дата заказа', tovar_id as 'Код товара', sklad_id as 'Склад', naimenovanie_tovara as 'Наименование товара', ed_izmer_tov as 'Ед. измер.', stoim_ed_tov as 'Цена', Zakaz.kolichestvo_tovara as 'Кол-во', Zakaz.stoimost_zakaza as 'Сумма', Klient.naimenovanie_klienta 'Клиент', Klient.inn_klienta as 'ИНН', Klient.telefon as 'Номер телефона' from tovar_sklad, tovar, sklad, zakaz, klient where tovar_sklad.sklad_id=sklad.id_sklada and tovar_sklad.tovar_id=tovar.id_tovara and Zakaz.tovar_sklad_id=Tovar_sklad.id_tovar_sklad and Zakaz.klient_id=Klient.id_klienta and klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + inn_klient + "');";
            MySqlConnection conn = DBUtils.GetDBConnection();
            MySqlDataAdapter sda = new MySqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                dataGridView4.DataSource = dt;
                dataGridView4.ClearSelection();
                sda.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.ClearSelection();
                this.dataGridView4.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView4.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[0].Width = 70;
                this.dataGridView4.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[1].Width = 90;
                this.dataGridView4.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[2].Width = 70;
                this.dataGridView4.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[3].Width = 70;
                this.dataGridView4.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[4].Width = 270;
                this.dataGridView4.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[5].Width = 50;
                this.dataGridView4.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[6].Width = 100;
                this.dataGridView4.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[7].Width = 100;
                this.dataGridView4.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[8].Width = 100;
                this.dataGridView4.Columns[9].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[9].Width = 250;
                this.dataGridView4.Columns[10].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[10].Width = 110;
                this.dataGridView4.Columns[11].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView4.Columns[11].Width = 110;
                dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn.Close();

                decimal Total = 0;

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    Total += Convert.ToDecimal(dataGridView4.Rows[i].Cells[8].Value);
                }

                label20.Text = "ИТОГО: " + Total.ToString("f2");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }
        }

        // Кнопка "Печать выборки заказов".
        private void button15_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString();
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                title.Range.Font.Size = 12;
                title.Range.Font.Color = Word.WdColor.wdColorDarkRed;
                title.Range.Text = $"Оптовая база \nВыборка заказов \n(номер заказа, дата заказа, код товара, склад, наименование, ед. изм., цена, кол-во, сумма, заказчик, ИНН, телефон)" +
                    $"\nДата формирования запроса: {date}";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView4.RowCount, dataGridView4.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 7;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView4.RowCount; i++)
                    for (int j = 1; j <= dataGridView2.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView4.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
                var footer = doc.Content.Paragraphs.Add();
                footer.Range.Font.Size = 14;
                footer.Range.Font.Color = Word.WdColor.wdColorDarkBlue;
                footer.Range.Text = $"{label20.Text}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // ******************* ВКЛАДКА "ОПАЛАТЫ" *********************
        //                Таблица "ОПЛАТЫ КЛИЕНТОВ"
        // Вывод данных в текстовые поля из dataGridView1 (Оплаты клиентов).
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox13.Text = dataGridView3.CurrentRow.Cells[0].Value.ToString();
            dateTimePicker2.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            textBox16.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
            comboBox4.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            textBox18.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
            maskedTextBox3.Text = dataGridView3.CurrentRow.Cells[5].Value.ToString();
        }

        // Автоматическое заполнение полей "ИНН" и "Номер телефона" при выборе клиента в выпадающем списке.
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string klient = comboBox4.Text;
            // ИНН.
            try
            {
                string INN = "select inn_klienta from Klient where naimenovanie_klienta ='" + klient + "';";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(INN, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(INN, conn);
                textBox18.Text = command.ExecuteScalar().ToString();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Номер телефона.
            try
            {
                string telefon = "select telefon from Klient where naimenovanie_klienta ='" + klient + "';";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(telefon, conn);

                conn.Open();
                MySqlCommand command = new MySqlCommand(telefon, conn);
                maskedTextBox3.Text = command.ExecuteScalar().ToString();
                conn.Close();
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Кнопка "Найти оплату" в таблице "Оплаты клиентов".
        private void button12_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы была заполнена строка поиска.
            if (textBox12.Text == null || textBox12.Text == "")
                MessageBox.Show(
                    "Вы не ввели запрос в строке поиска!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {

                for (int i = 0; i < dataGridView3.RowCount; i++)
                {
                    dataGridView3.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView3.ColumnCount; j++)
                        if (dataGridView3.Rows[i].Cells[j].Value != null)
                            if (dataGridView3.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox12.Text.ToLower()))
                            {
                                dataGridView3.Rows[i].Selected = true;
                                break;
                            }
                }
            }
            textBox1.Clear();
        }

        // Кнопка "Печать оплат" из таблицы "Оплаты клиентов".
        private void button11_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString();
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                title.Range.Font.Color = Word.WdColor.wdColorDarkRed;
                title.Range.Font.Size = 14;
                title.Range.Text = $"Оптовая база\nОплаты клиентов (номер документа, дата оплаты, сумма оплаты, наименование клиента, ИНН, номер телефона)" +
                    $"\nДата формирования списка: {date}";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView3.RowCount, dataGridView3.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 12;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView3.RowCount; i++)
                    for (int j = 1; j <= dataGridView3.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView3.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
                var footer = doc.Content.Paragraphs.Add();
                footer.Range.Font.Size = 14;
                footer.Range.Font.Color = Word.WdColor.wdColorDarkBlue;
                footer.Range.Text = $"{itog}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // Кнопка "Добавить оплату".
        private void button14_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox4.Text.Equals("") || textBox16.Text == null || textBox16.Text == "")
            {
                MessageBox.Show(
                "Проверьте введенные данные!",
                "Сообщение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
                return;
            }
            else
            {
                string klient = comboBox4.Text;
                string data = dateTimePicker2.Value.ToString("yyyy-MM-dd");
                var oplata = textBox16.Text.Replace(",", ".");
                DialogResult res = MessageBox.Show($"Провести оплату от клиента:\n\n{klient}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = "INSERT INTO Oplata SET klient_id=(select klient.id_klienta from klient where klient.naimenovanie_klienta='" + klient + "'), data_oplaty='" + data + "', summa_oplaty='" + oplata + "'; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    textBox2.Clear();
                    textBox3.Clear();
                    maskedTextBox1.Clear();
                }
            }
        }

        // Кнопка "Изменить оплату".
        private void button13_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода/вывода данных.
            if (comboBox4.Text.Equals("") || textBox13.Text == null || textBox13.Text == "" || textBox16.Text == null || textBox16.Text == "" || textBox18.Text == null || textBox18.Text == "")
            {
                MessageBox.Show(
                    "Не заполнены поля ввода.",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            else
            {
                string name = dataGridView3.CurrentRow.Cells[3].Value.ToString();
                DialogResult res = MessageBox.Show($"Изменить данные оплаты для клиента:\n\n{name}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    int document = int.Parse(dataGridView3.CurrentRow.Cells[0].Value.ToString());
                    string data = dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    var oplata = textBox16.Text.Replace(",", ".");
                    string query = " UPDATE oplata SET klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + textBox18.Text + "'), data_oplaty='" + data + "', summa_oplaty='" + oplata + "' WHERE id_dokumenta=" + document + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                    textBox2.Clear();
                    textBox3.Clear();
                    maskedTextBox1.Clear();
                }
            }
        }

        // Кнопка "Удалить оплату".
        private void button16_Click(object sender, EventArgs e)
        {
            string name = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            DialogResult res = MessageBox.Show($"Удалить оплату клиента: \n\n{name}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                string valueCell = dataGridView3.CurrentCell.Value != null ? dataGridView3.CurrentCell.Value.ToString() : "";
                string del = "delete from oplata where id_dokumenta = " + valueCell + ";";
                Action(del);
                GetInfo(ID);
            }
            else
            {
                MessageBox.Show("Не выбрано ни одной записи! Удаление невозможно.");
            }
        }

        // Кнопка "Поиск индивидуальной задолженности клиента".
        private void button18_Click(object sender, EventArgs e)
        {
            this.label26.ForeColor = System.Drawing.Color.Navy;
            dataGridView5.AllowUserToAddRows = false;
            // Проверяем, чтобы были заполнены поля ввода.
            if (textBox18.Text == null || textBox18.Text == "")
            {
                MessageBox.Show(
                    "Для просмотра задолженности необходимо выбрать клиента!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            else
            {
                label26.Text = comboBox4.Text;
                string query = "SELECT (select sum(Zakaz.stoimost_zakaza) from Zakaz where Zakaz.klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + textBox18.Text + "')) AS 'Дебет', (select sum(Oplata.summa_oplaty) from Oplata where Oplata.klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + textBox18.Text + "')) AS 'Кредит', ((select sum(Oplata.summa_oplaty) from Oplata where Oplata.klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + textBox18.Text + "')) - (select sum(Zakaz.stoimost_zakaza) from Zakaz where Zakaz.klient_id=(select klient.id_klienta from klient where klient.inn_klienta='" + textBox18.Text + "'))) as 'Сальдо';";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlDataAdapter sda = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                try
                {
                    conn.Open();
                    dataGridView5.DataSource = dt;
                    dataGridView5.ClearSelection();
                    sda.Fill(dt);
                    dataGridView5.DataSource = dt;
                    dataGridView5.ClearSelection();
                    this.dataGridView5.ForeColor = System.Drawing.Color.Navy;
                    this.dataGridView5.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView5.Columns[0].Width = 90;
                    this.dataGridView5.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView5.Columns[1].Width = 90;
                    this.dataGridView5.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    this.dataGridView5.Columns[2].Width = 90;
                    dataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
                }
            }
        }

        // Просмотр суммы общей задолженности клиентов.
        private void button17_Click(object sender, EventArgs e)
        {
            string query = "SELECT (select sum(Zakaz.stoimost_zakaza) from Zakaz) AS 'Дебет', (select sum(Oplata.summa_oplaty) from Oplata) AS 'Кредит', ((select sum(Oplata.summa_oplaty) from Oplata) - (select sum(Zakaz.stoimost_zakaza) from Zakaz)) as 'Сальдо';";
            MySqlConnection conn = DBUtils.GetDBConnection();
            MySqlDataAdapter sda = new MySqlDataAdapter(query, conn);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                dataGridView6.DataSource = dt;
                dataGridView6.ClearSelection();
                sda.Fill(dt);
                dataGridView6.DataSource = dt;
                dataGridView6.ClearSelection();
                this.dataGridView6.ForeColor = System.Drawing.Color.Navy;
                this.dataGridView6.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView6.Columns[0].Width = 90;
                this.dataGridView6.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView6.Columns[1].Width = 90;
                this.dataGridView6.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView6.Columns[2].Width = 90;
                dataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла непредвиденая ошибка!" + Environment.NewLine + ex.Message);
            }
        }

        // Запрет на ввод в поле "Сумма оплаты" любых символов и букв, кроме чисел, запятой и клавиши backspace.
        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (e.KeyChar == Convert.ToChar(",")) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        // ******************* ВКЛАДКА "ТОВАРНАЯ НОМЕНКЛАТУРА" *********************
        //                     Таблица "ТОВАРНАЯ НОМЕНКЛАТУРА"
        // Вывод данных в текстовые поля из dataGridView7 (Товарная номенклатура).
        private void dataGridView7_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox14.Text = dataGridView7.CurrentRow.Cells[0].Value.ToString();
            this.textBox14.ForeColor = System.Drawing.Color.Navy;
            comboBox3.Text = dataGridView7.CurrentRow.Cells[1].Value.ToString();
            this.comboBox3.ForeColor = System.Drawing.Color.Navy;
            comboBox5.Text = dataGridView7.CurrentRow.Cells[2].Value.ToString();
            this.comboBox5.ForeColor = System.Drawing.Color.Navy;
            textBox21.Text = dataGridView7.CurrentRow.Cells[3].Value.ToString();
            this.textBox21.ForeColor = System.Drawing.Color.Navy;
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        // Вывод данных в текстовые поля из dataGridView7 (Товарная номенклатура).
        private void dataGridView9_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox15.Text = dataGridView9.CurrentRow.Cells[0].Value.ToString();
            this.textBox15.ForeColor = System.Drawing.Color.Navy;
            comboBox8.Text = dataGridView9.CurrentRow.Cells[2].Value.ToString();
            this.comboBox8.ForeColor = System.Drawing.Color.Navy;
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        // Кнопка "Поиск в таблице - Товарная номенклатура".
        private void button20_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы была заполнена строка поиска.
            if (textBox17.Text == null || textBox17.Text == "")
                MessageBox.Show(
                    "Вы не ввели запрос в строке поиска!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {

                for (int i = 0; i < dataGridView7.RowCount; i++)
                {
                    dataGridView7.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView7.ColumnCount; j++)
                        if (dataGridView7.Rows[i].Cells[j].Value != null)
                            if (dataGridView7.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox17.Text.ToLower()))
                            {
                                dataGridView7.Rows[i].Selected = true;
                                break;
                            }
                }
            }
            textBox17.Clear();
        }

        // Кнопка "Печть товарной номенклатуры".
        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                DateTime dt = DateTime.Now;
                string curDate = dt.ToShortDateString();
                title.Range.Font.Size = 14;
                title.Range.Text = $"Оптовая база\nПрайс-лист \n(код товара, склад, наименование товара, ед. измер., цена)" +
                    $"\nДата формирования списка: {curDate}";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView7.RowCount, dataGridView7.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 7;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView7.RowCount; i++)
                    for (int j = 1; j <= dataGridView7.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView7.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // Запрет на ввод в поле "Цена" любых символов и букв, кроме чисел, запятой и клавиши backspace.
        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (e.KeyChar == Convert.ToChar(",")) | e.KeyChar == '\b') return;
            else
                e.Handled = true;
        }

        // Кнопка "Добавить товар" во вкладке "Склады".
        private void button22_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox3.Text.Equals("") || comboBox5.Text.Equals("") || textBox21.Text == null || textBox21.Text == "")
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                DialogResult res = MessageBox.Show($"Добавить товар:\n\n{comboBox3.Text}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    var stoimost = textBox21.Text.Replace(",", ".");
                    string query = "INSERT INTO tovar SET naimenovanie_tovara='" + comboBox3.Text + "', ed_izmer_tov='" + comboBox5.Text + "', stoim_ed_tov='" + stoimost + "'; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Кнопка "Изменить товар" во вкладке "Товарная номенклатура".
        private void button21_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox3.Text.Equals("") || comboBox5.Text.Equals("") || textBox21.Text == null || textBox21.Text == "")
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                string name = dataGridView7.CurrentRow.Cells[1].Value.ToString();
                DialogResult res = MessageBox.Show($"Изменить данные товара:\n\n{name}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    int id_tov = int.Parse(dataGridView7.CurrentRow.Cells[0].Value.ToString());
                    var stoimost = textBox21.Text.Replace(",", ".");
                    string query = " UPDATE tovar SET naimenovanie_tovara='" + comboBox3.Text + "', ed_izmer_tov='" + comboBox5.Text + "', stoim_ed_tov='" + stoimost + "' WHERE id_tovara=" + id_tov + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Перемещение товара из списка номенклатуры в ассортиментный перечень склада.
        private void button29_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox8.Text.Equals("") || textBox15.Text == null || textBox15.Text == "")
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                string sklad = comboBox8.Text;
                DialogResult res = MessageBox.Show($"Добавить товар на \n\n{sklad}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = " INSERT INTO tovar_sklad (sklad_id, tovar_id) VALUES ((select sklad.id_sklada from sklad where sklad.naimenovanie_sklada='" + comboBox8.Text + "'), '" + textBox15.Text + "'); ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Кнопка "Изменить расположения товара на складах".
        private void button25_Click(object sender, EventArgs e)
        {
            textBox15.Enabled = false;
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox8.Text.Equals("") || textBox15.Text == null || textBox15.Text == "")
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                int id_tovar_sklad = int.Parse(dataGridView9.CurrentRow.Cells[0].Value.ToString());
                DialogResult res = MessageBox.Show($"Переместить товар с кодом {textBox15.Text}\nна {comboBox8.Text}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = " UPDATE tovar_sklad SET sklad_id=(select sklad.id_sklada from sklad where sklad.naimenovanie_sklada='" + comboBox8.Text + "'), tovar_id='" + textBox15.Text + "' WHERE id_tovar_sklad=" + id_tovar_sklad + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Поиск в таблице "Товары на складе".
        private void button28_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы была заполнена строка поиска.
            if (textBox19.Text == null || textBox19.Text == "")
                MessageBox.Show(
                    "Вы не ввели запрос в строке поиска!",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            else
            {

                for (int i = 0; i < dataGridView9.RowCount; i++)
                {
                    dataGridView9.Rows[i].Selected = false;
                    for (int j = 0; j < dataGridView9.ColumnCount; j++)
                        if (dataGridView9.Rows[i].Cells[j].Value != null)
                            if (dataGridView9.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox19.Text.ToLower()))
                            {
                                dataGridView9.Rows[i].Selected = true;
                                break;
                            }
                }
            }
            textBox19.Clear();
        }

        // Кнопка "Печать списка товаров на складах".
        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                Word._Application app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var title = doc.Content.Paragraphs.Add();
                DateTime dt = DateTime.Now;
                string curDate = dt.ToShortDateString();
                title.Range.Font.Color = Word.WdColor.wdColorDarkRed;
                title.Range.Font.Size = 14;
                title.Range.Text = $"Оптовая база\nСписок товаров на складах \n(код товара, наименование товара, склад)" +
                    $"\nДата формирования списка: {curDate}";
                title.Range.Font.Name = "Arial Narrow";
                object objMissing = System.Reflection.Missing.Value;
                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, dataGridView9.RowCount, dataGridView9.ColumnCount, ref objMissing, ref objMissing);
                table.Range.Paragraphs.SpaceAfter = 6;
                table.Range.Font.Name = "Arial Narrow";
                table.Range.Font.Size = 9;
                // Заполнение ячеек таблицы.
                for (int i = 1; i < dataGridView9.RowCount; i++)
                    for (int j = 1; j <= dataGridView9.ColumnCount; j++)
                    {
                        {
                            var cell = dataGridView9.Rows[i - 1].Cells[j - 1];
                            table.Cell(i, j).Range.Text = cell.Value.ToString();
                        }
                    }
                table.Borders.Enable = 3;
                table.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка!");
                MessageBox.Show(ex.Message);
            }
        }

        // ******************* ВКЛАДКА "СКЛАДЫ" *********************
        //                     Таблица "СКЛАДЫ"
        // Вывод данных в текстовые поля из dataGridView8 (Склады).
        private void dataGridView8_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBox6.Text = dataGridView8.CurrentRow.Cells[1].Value.ToString();
            this.comboBox6.ForeColor = System.Drawing.Color.Navy;
            comboBox7.Text = dataGridView8.CurrentRow.Cells[2].Value.ToString();
            this.comboBox5.ForeColor = System.Drawing.Color.Navy;
            maskedTextBox4.Text = dataGridView8.CurrentRow.Cells[3].Value.ToString();
            this.textBox21.ForeColor = System.Drawing.Color.Navy;
        }

        // Кнопка "Добавить склад".
        private void button24_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox6.Text.Equals("") || comboBox7.Text.Equals(""))
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                DialogResult res = MessageBox.Show($"Добавить склад:\n\n{comboBox6.Text}?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = "INSERT INTO sklad (naimenovanie_sklada, fio_zavsklada, telefon) VALUES ('" + comboBox6.Text + "', '" + comboBox7.Text + "', '" + maskedTextBox4.Text + "'); ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Кнопка "Изменить данные склада".
        private void button23_Click(object sender, EventArgs e)
        {
            // Проверяем, чтобы были заполнены поля ввода.
            if (comboBox6.Text.Equals("") || comboBox7.Text.Equals(""))
            {
                MessageBox.Show("Введите данные.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                int id_sklad = int.Parse(dataGridView8.CurrentRow.Cells[0].Value.ToString());
                DialogResult res = MessageBox.Show($"Внести изменения:\n\n{comboBox6.Text}\n\nЗавсклад: {comboBox7.Text}\n\nНомер телефона: {maskedTextBox4.Text} ?", "Подтвердите действие", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    string query = " UPDATE sklad SET naimenovanie_sklada='" + comboBox6.Text + "', fio_zavsklada='" + comboBox7.Text + "', telefon='" + maskedTextBox4.Text + "' WHERE id_sklada=" + id_sklad + "; ";
                    MySqlConnection conn = DBUtils.GetDBConnection();
                    MySqlCommand cmDB = new MySqlCommand(query, conn);
                    try
                    {
                        conn.Open();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                    }
                    Action(query);
                    GetInfo(ID);
                }
            }
        }

        // Кнопка "Изменить логин и пароль".
        private void button26_Click(object sender, EventArgs e)
        {
            if (button26.Text == "Изменить")
            {
                textBox20.Visible = true;
                textBox22.Visible = true;
                button26.Text = "Сохранить";
            }
            else if (button26.Text == "Сохранить")
            {
                string query = "update avtorizacia set login ='" + textBox20.Text + "', password ='" + textBox22.Text + "' where id_user = " + ID.ToString() + ";";
                MySqlConnection conn = DBUtils.GetDBConnection();
                MySqlCommand cmDB = new MySqlCommand(query, conn);
                try
                {
                    conn.Open();
                    cmDB.ExecuteReader();
                    conn.Close();
                    textBox20.Visible = false;
                    textBox22.Visible = false;
                    button26.Text = "Изменить";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла непредвиденная ошибка!" + Environment.NewLine + ex.Message);
                }
            }
        }
    }
}
