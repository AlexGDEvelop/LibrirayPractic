﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;

namespace Libriray
{
    public partial class main : Form
    {
        public static string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=er.mdb;";
        private OleDbConnection conn;
        private OleDbDataAdapter myDataAdapter;
        private DataSet myDataSet;
        private DataTable myDt;
        BindingSource bs1 = new BindingSource();
        ComboBox comboBox;
        string[] names;
        Control tb;
        //string ds[];

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

        Dictionary<string, string> comboQuerys = new Dictionary<string, string>
        {
            {"pub","" },
            {"give_pub","" },
            {"return_pub","" },
            {"rcpnt","" },
            {"num_of_pub","" },
            {"pub_spec","" },
            {"type_of_pub","" },
            {"discipline","" },
            {"author","" },
            {"pstions","" }
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
            {"pub_spec",new string[]{  "SELECT name FROM type_of_pub"  } },
            {"type_of_pub",new string[]{ } },
            {"discipline",new string[]{ }},
            {"author",new string[]{ }},
            {"pstions", new string[]{ } }
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


        Dictionary<string, string> insertQuerys = new Dictionary<string, string>
        {
            {"pub","" },
            {"give_pub","" },
            {"return_pub","" },
            {"rcpnt","" },
            {"num_of_pub","" },
            {"pub_spec","" },
            {"type_of_pub","" },
            {"discipline","" },
            {"author","" },
            {"pstions","" }
        };

        string[] tabNames = new string[10] { "pub","give_pub","return_pub","rcpnt","num_of_pub","pub_spec","type_of_pub","discipline","author","pstions" };



        public main()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        private void main_Load(object sender, EventArgs e)
        {

            //bitmap1 = Bitmap.FromResource();

            conn = new OleDbConnection(connectionString);
            myDataAdapter = new OleDbDataAdapter();

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
            //myDataSet.Tables[0].Columns[0].ColumnName;
            string[] qs = new string[myDataSet.Tables[0].Columns.Count];
            foreach (DataColumn col in myDataSet.Tables[0].Columns)
            {
                query += col.ColumnName + ",";
            }

            foreach (var tb in panel1.Controls)
            {
                //qs[panel1.Controls.GetChildIndex(tb)/2] = tb.Text;
                Console.WriteLine(tb);
            }

            /*foreach (ComboBox combo in panel1.Controls)
            {
                qs[panel1.Controls.GetChildIndex(combo)/2] = combo.SelectedIndex.ToString();
            }*/

            query = query.Remove(query.Length - 1);

            query += ") VALUES (";


            for (int i = 0; i < qs.Length; i++)
            {
                query += qs[i] + ",";
            }

            query = query.Remove(query.Length - 1);

            query += ")";


            Console.WriteLine(query);


        }

        private void deystv_combo_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (deystv_combo.SelectedIndex == 1)
            {
                panel1.Visible = true;
                //Добавление записей на панель

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
                        myDataSet = new DataSet();
                        myDt = new DataTable();
                        myDataAdapter.Fill(myDataSet,"dt");
                        myDataAdapter.Update(myDataSet, "dt");
                        //tb.DataBindings.Clear();

                        string[] ds = new string[myDataSet.Tables[0].Rows.Count];
                        BindingList<string> ts =new BindingList<string>();
                        for (int f = 0; f < myDataSet.Tables[0].Rows.Count; f++)
                        {

                            ts.Insert(f, myDataSet.Tables[0].Rows[f][0].ToString());
                            
                            

                        }
                        //Binding d = new Binding("DataSource", myDataSet, "dt." + combo_querys[tabNames[tabs_combo.SelectedIndex]][1, j], true);
                        Binding d = new Binding("DataSource",ts, "",true);
                        //Binding gf = new Binding("DisplayMember", "", "", true);
                        comboBox.DataSource = ts;
                       // tb.DataBindings.Add(gf);
                        //comboBox1.DataSource = ts;

                       // comboBox1.DisplayMember = ts;

                        conn.Close();
                        j++;
                        panel1.Controls.Add(comboBox);
                    }

                    panel1.Controls.Add(tb);


                    panel1.Controls.Add(lbl);

                    
                    /*try
                    {
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        
                    }*/
                   

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
                //query += myDataSet.Tables[0].Columns[i].ColumnName + " LIKE " + search_box_tb.Text +"% OR ";
            }


            bs1.Filter = query;
        }

        
    }  
}
