using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;

using Excel = Microsoft.Office.Interop.Excel;

namespace Libriray
{
    public partial class main : Form
    {
        public static string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=er.mdb;";
        private OleDbConnection conn;
        private OleDbDataAdapter myDataAdapter,dA;
        private DataSet myDataSet,comboDataSet,dataSet;
        string[] qs;
        private DataTable myDt;
        BindingSource bs1 = new BindingSource();
        ComboBox comboBox;
        string[] names;
        Control tb;
        const int Margin = 26;
        const int tbstart_pos = 15;
        int tb_pos = 0;
        const int lb_start_pos = 18;
        int lb_pos = 0;
        const int lb_x_pos = 12;
        const int tb_x_pos = 128;
        OleDbCommand myOleDbCommand;
        Dictionary<string, string> selecQuerys = new Dictionary<string, string>
        {
            { "pub", "SELECT pub.name AS [Название издания], author.name AS [Имя автора], author.last_n AS [Фамилия автора], pub_spec.name AS [Вид издания], discipline.name AS [Дисциплина] " +
                     "FROM discipline INNER JOIN (pub_spec INNER JOIN (author INNER JOIN pub ON author.id_author = pub.id_author) ON pub_spec.id_pub_spec = pub.id_pub_spec) ON discipline.id_discipline = pub.id_discipline;" },

            { "give_pub", "SELECT  rcpnt.last_n AS [Фамилия получателя],rcpnt.name AS [Имя получатяеля], pstions.name AS [Должность], pub.name AS [Издание] , give_pub.num AS [Количество], give_pub.datetime AS [Дата] " +
                " FROM pstions INNER JOIN (pub INNER JOIN (rcpnt INNER JOIN give_pub ON rcpnt.id_rcpnt = give_pub.id_rcpnt) ON pub.id_pub = give_pub.id_pub) ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "return_pub", "SELECT rcpnt.last_n AS [Фамилия получателя], rcpnt.name  AS [Имя получателя], pstions.name  AS [Должность], pub.name  AS [Издание], return_pub.num AS [Количество], return_pub.datetime  AS [Дата] " +
                " FROM pstions INNER JOIN (pub INNER JOIN (rcpnt INNER JOIN return_pub ON rcpnt.id_rcpnt = return_pub.id_rcpnt) ON pub.id_pub = return_pub.id_pub) ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "rcpnt", "SELECT rcpnt.last_n AS [Фамилия], rcpnt.name AS [Имя], rcpnt.mid_n AS [Отчество], pstions.name AS [Должность]" +
                " FROM pstions INNER JOIN rcpnt ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "num_of_pub", "SELECT pub.name AS [Издание], num_of_pub.num AS [Количество]" +
                " FROM pub INNER JOIN num_of_pub ON pub.id_pub = num_of_pub.id_pub;" },
            { "pub_spec", "SELECT pub_spec.name AS [Вид издания], type_of_pub.name AS [Тип издания]" +
                " FROM type_of_pub INNER JOIN pub_spec ON type_of_pub.id_type_of_pub = pub_spec.id_type_of_pub;" },
            { "type_of_pub", "SELECT type_of_pub.name AS [Тип издания]" +
                " FROM type_of_pub;" },
            { "discipline", "SELECT discipline.name AS [Дисциплина]" +
                " FROM discipline;" },
            { "author", "SELECT author.last_n AS [Фамилия], author.name AS [Имя], author.mid_n AS [Отчество]" +
                " FROM author;" },
            { "pstions", "SELECT pstions.name AS [Должность]" +
                " FROM pstions;" },
        };

        Dictionary<string, Control[]> insertElements = new Dictionary<string, Control[]>
        {
            {"pub", new Control[] { new TextBox() {}, new ComboBox() {Name = "c" }, new ComboBox() { Name = "c" }, new ComboBox() { Name = "c" } } },
            {"give_pub", new Control[] { new ComboBox() { Name = "c" }, new ComboBox() { Name = "c" }, new ComboBox() { Name = "c" }, new TextBox(), new DateTimePicker() } },
            {"return_pub", new Control[] { new ComboBox() { Name = "c" }, new ComboBox() { Name = "c" }, new ComboBox() { Name = "c" }, new TextBox(), new DateTimePicker()} },
            {"rcpnt", new Control[] { new TextBox(), new TextBox() { }, new TextBox(), new ComboBox() { Name = "c" } } },
            {"num_of_pub", new Control[] { new ComboBox() { Name = "c" }, new TextBox()} },
            {"pub_spec", new Control[] { new TextBox(), new ComboBox() { Name = "c" } } },
            {"type_of_pub", new Control[] { new TextBox() } },
            {"discipline", new Control[] { new TextBox()} },
            {"author", new Control[] { new TextBox(), new TextBox() {}, new TextBox() } },
            {"pstions", new Control[] { new TextBox()} }
        };

        Dictionary<string, string[]>  combo_querys= new Dictionary<string, string[]>
        {
            {"pub",new string[]{ "SELECT last_n + name AS [snam] FROM author", "SELECT DISTINCT name FROM pub_spec", "SELECT DISTINCT name FROM discipline"  } },
            {"give_pub",new string[]{ "SELECT last_n + name AS [snam] FROM rcpnt", "SELECT name FROM pstions","SELECT name FROM pub"} },
            {"return_pub",new string[]{ "SELECT last_n + name AS [snam] FROM rcpnt", "SELECT name FROM pstions", "SELECT name FROM pub" } },
            {"rcpnt",new string[]{ "SELECT name FROM pstions" } },
            {"num_of_pub",new string[]{  "SELECT name FROM pub"  } },
            {"pub_spec",new string[]{  "SELECT name FROM type_of_pub"  } }
        };
        


        Dictionary<string, string[]> tabsCaption = new Dictionary<string, string[]>
        {
            {"pub", new string[] { "Название издания", "Автор", "Вид издания", "Дисциплина" } },
            {"give_pub",new string[] { "Получатель", "Должность", "Издание", "Количество", "Дата"}},
            {"return_pub",new string[] { "Получатель", "Должность", "Издание", "Количество", "Дата"} },
            {"rcpnt",new string[] { "Фамилия", "Имя","Отчество", "Должность"}},
            {"num_of_pub",new string[] {"Издание","Количество"}},
            {"pub_spec",new string[] { "Вид издания", "Тип издания"}},
            {"type_of_pub",new string[] { "Тип издания"} },
            {"discipline",new string[] { "Дисциплина" }},
            {"author",new string[] { "Фамилия", "Имя","Отчество"}},
            {"pstions",new string[] { "Должность"}}
        };

        string[] tabNames = new string[10] { "pub","give_pub","return_pub","rcpnt","num_of_pub","pub_spec","type_of_pub","discipline","author","pstions" };



        public main()
        {
            InitializeComponent();
        }

        private void main_Load(object sender, EventArgs e)
        {

            //bitmap1 = Bitmap.FromResource();

            conn = new OleDbConnection(connectionString);
            myDataAdapter = new OleDbDataAdapter();
            dA = new OleDbDataAdapter();

            panel1.Visible = false;

            tabs_combo.SelectedIndex = 0;

            ViewTab("pub");

            
        }


        void ViewTab(string who)
        {


            conn.Open();
            myOleDbCommand = new OleDbCommand(selecQuerys[who], conn);
            myDataAdapter.SelectCommand = myOleDbCommand;
            myDataSet = new DataSet();
            myDataAdapter.Fill(myDataSet);
            bs1 = new BindingSource(myDataSet, myDataSet.Tables[0].ToString());
            dataGridView.DataSource = bs1;
            myDataAdapter.Update(myDataSet);
            qs = new string[myDataSet.Tables[0].Columns.Count];
            dgvnav.BindingSource = bs1;

            conn.Close();
        }

        private void tabs_combo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (deystv_combo.SelectedIndex == 1)
            {
                deystv_combo_SelectedIndexChanged(sender, e);
            }
            ViewTab(tabNames[tabs_combo.SelectedIndex]);
        }


        void InsertInTab(string who)
        {
            string query = "INSERT INTO " + who + " (";



            conn.Open();
            myOleDbCommand = new OleDbCommand("SELECT * FROM " + who, conn);
            dA.SelectCommand = myOleDbCommand;
            dataSet = new DataSet();
            dA.Fill(dataSet);
            conn.Close();

            foreach (DataColumn col in dataSet.Tables[0].Columns)
            {
                query += col.ColumnName + ",";
            }

            ComboBox[] combo = panel1.Controls.OfType<ComboBox>().ToArray();
            TextBox[] tb = panel1.Controls.OfType<TextBox>().ToArray();


            query = query.Remove(query.Length - 1);

            query += ") VALUES (";


            foreach(ComboBox comb in combo)
            {
                if (comb.Items.Count != 0)
                {
                    qs[int.Parse(comb.Name.Substring(comb.Name.Length - 1))] = (comb.SelectedIndex + 1).ToString();
                }
                else
                {
                    MessageBox.Show("В одной из таблиц подстановки нет значений","Внимание");
                    return;
                }

            }

            foreach(TextBox text in tb)
            {
                if (text.Text != "")
                {
                    qs[int.Parse(text.Name.Substring(text.Name.Length - 1))] = "'" + text.Text + "'";
                }
                else
                {
                    MessageBox.Show("Все поля должны быть заполнены", "Внимане");
                }
                
            }

            query += (dataSet.Tables[0].Rows.Count + 1).ToString() + ",";
            for (int i = 0; i < qs.Length; i++)
            {
                query += qs[i] + ",";
            }

            query = query.Remove(query.Length - 1);

            query += ")";


            conn.Open();
            myOleDbCommand = new OleDbCommand(query, conn);
            myDataAdapter.SelectCommand = myOleDbCommand;
            dataSet = new DataSet();
            myDataAdapter.Fill(dataSet);
            myDataAdapter.Update(myDataSet);
            conn.Close();

            ViewTab(who);


            Console.WriteLine(query);


        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {

            conn.Open();
            myOleDbCommand = new OleDbCommand("SELECT * FROM " + tabNames[tabs_combo.SelectedIndex], conn);
            dA.SelectCommand = myOleDbCommand;
            dataSet = new DataSet();
            dA.Fill(dataSet);
            conn.Close();

            string query = string.Format("DELETE FROM {0} WHERE {1}={2}", tabNames[tabs_combo.SelectedIndex],dataSet.Tables[0].Columns[0].ColumnName,dataGridView.CurrentRow.Index + 2);

            conn.Open();
            myOleDbCommand = new OleDbCommand(query, conn);
            myDataAdapter.DeleteCommand = myOleDbCommand;
            dataSet = new DataSet();

            myDataAdapter.Fill(dataSet);

            try
            {
                myDataAdapter.Update(myDataSet);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Ошибка");

                
            }

            conn.Close();

            ViewTab(tabNames[tabs_combo.SelectedIndex]);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            OtchGen gen = new OtchGen();
            gen.Otchgen(dataGridView, tabNames[tabs_combo.SelectedIndex]);
        }

        private void deystv_combo_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (deystv_combo.SelectedIndex == 1)
            {
                panel1.Visible = true;
                lb_pos = lb_start_pos;
                tb_pos = tbstart_pos;
                panel1.Controls.Clear();

                bindingNavigatorAddNewItem.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                toolStripButton1.Enabled = true;

                int j = 0;
                for (int i = 0; i < tabsCaption[tabNames[tabs_combo.SelectedIndex]].Length; i++)
                {
                    Label lbl = new Label()
                    {
                        Name = "lbl" + i,
                        Location = new Point(lb_x_pos, lb_pos),
                        Text = tabsCaption[tabNames[tabs_combo.SelectedIndex]][i]
                    };

                    try
                    {
                        tb = insertElements[tabNames[tabs_combo.SelectedIndex]][i];
                        tb.Name += "tb" + i;
                        tb.Location = new Point(tb_x_pos, tb_pos);
                    }
                    catch
                    {

                    }

                    if (tb.Name.StartsWith("c"))
                    {


                        comboBox = new ComboBox() {  Name = tb.Name, Location = tb.Location };
                        tb = null;
                        conn.Open();
                        myOleDbCommand = new OleDbCommand(combo_querys[tabNames[tabs_combo.SelectedIndex]][j], conn);
                        myDataAdapter.SelectCommand = myOleDbCommand;
                        comboDataSet = new DataSet();
                        myDt = new DataTable();
                        myDataAdapter.Fill(comboDataSet, "dt");
                        myDataAdapter.Update(comboDataSet, "dt");

                        string[] ds = new string[comboDataSet.Tables[0].Rows.Count];
                        BindingList<string> ts =new BindingList<string>();
                        for (int f = 0; f < comboDataSet.Tables[0].Rows.Count; f++)
                        {

                            ts.Insert(f, comboDataSet.Tables[0].Rows[f][0].ToString());
                            
                            

                        }

                        Binding d = new Binding("DataSource",ts, "",true);

                        comboBox.DataSource = ts;


                        conn.Close();
                        j++;
                        panel1.Controls.Add(comboBox);
                    }

                    panel1.Controls.Add(tb);

                    panel1.Controls.Add(lbl);

                    lb_pos += Margin;
                    tb_pos += Margin;
                }
                j = 0;
            }
            else
            {
                panel1.Visible = false;
                bindingNavigatorAddNewItem.Enabled = false;
                bindingNavigatorDeleteItem.Enabled = false;
                toolStripButton1.Enabled = false;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            InsertInTab(tabNames[tabs_combo.SelectedIndex]);
        }

        private void search_box_tb_TextChanged(object sender, EventArgs e)
        {
            string query = " ";

            names = new string[myDataSet.Tables[0].Columns.Count];
            for (int i = 0; i < names.Length; i++)
            {
                names[i] = myDataSet.Tables[0].Columns[i].ColumnName;

                if (names.Length == i + 1 | names.Length == 1)
                {
                    query += "["+myDataSet.Tables[0].Columns[i].ColumnName + "]" + " LIKE '" + search_box_tb.Text + "%' ";
                }
                else
                {
                    query += "[" + myDataSet.Tables[0].Columns[i].ColumnName + "]" + " LIKE '" + search_box_tb.Text + "%' OR ";
                }
            }

            bs1.Filter = query;
        }
    }  

    public class OtchGen
    {
        Excel.Application excelapp;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        public void Otchgen(DataGridView grid, string who)
        {

            MessageBox.Show("Производится сохранение, не выключайте программу.\n По завершению, появится сообщение 'Готово'", "Сохранение");
            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + "Save_" + who + "_" + DateTime.Now.ToShortDateString() + ".xlsx";

            ExcelStart();


            for (int i = 0; i < grid.RowCount; i++)
            {

                for (int j = 0; j < grid.ColumnCount; j++)
                {
                    worksheet.Rows[1].Columns[j + 1] = grid.Columns[j].Name;
                    worksheet.Rows[i + 2].Columns[j + 1] = grid.Rows[i].Cells[j].Value;

                }
            }
            ExcelSaveAndExit(path);

            MessageBox.Show("Готово", "Сохранено");

        }
        void ExcelStart()
        {

            excelapp = new Excel.Application();
            workbook = excelapp.Workbooks.Add();
            worksheet = workbook.ActiveSheet;
        }

        void ExcelSaveAndExit(string path)
        {
            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();

        }
    }
}
