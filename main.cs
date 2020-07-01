using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Libriray
{
    public partial class main : Form
    {
        public static string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=er.mdb;";
        private OleDbConnection conn;
        private OleDbDataAdapter myDataAdapter;
        private DataSet myDataSet;
        private BindingSource bindingSource;
        OleDbCommand myOleDbCommand;
        Dictionary<string, string> selecQuerys = new Dictionary<string, string> {
            { "pub", "SELECT pub.name AS [Название издания], author.name AS [Имя автора], author.last_n AS [Фамилия автора], pub_spec.name AS [Вид издания], discipline.name AS [Дисциплина] " +
                     "FROM discipline INNER JOIN (pub_spec INNER JOIN (author INNER JOIN pub ON author.id_author = pub.id_author) ON pub_spec.id_pub_spec = pub.id_pub_spec) ON discipline.id_discipline = pub.id_discipline;" },

            { "give_pub", "SELECT rcpnt.name, rcpnt.last_n, pstions.name, pub.name, give_pub.num, give_pub.datetime " +
                " FROM pstions INNER JOIN (pub INNER JOIN (rcpnt INNER JOIN give_pub ON rcpnt.id_rcpnt = give_pub.id_rcpnt) ON pub.id_pub = give_pub.id_pub) ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "return_pub", "SELECT rcpnt.last_n, rcpnt.name, pstions.name, pub.name, return_pub.num, return_pub.datetime " +
                " FROM pstions INNER JOIN (pub INNER JOIN (rcpnt INNER JOIN return_pub ON rcpnt.id_rcpnt = return_pub.id_rcpnt) ON pub.id_pub = return_pub.id_pub) ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "rcpnt", "SELECT rcpnt.last_n, rcpnt.name, rcpnt.mid_n, pstions.name" +
                " FROM pstions INNER JOIN rcpnt ON pstions.id_pstion = rcpnt.id_pstion;" },
            { "num_of_pub", "SELECT pub.name, num_of_pub.num" +
                " FROM pub INNER JOIN num_of_pub ON pub.id_pub = num_of_pub.id_pub;" },
            { "pub_spec", "SELECT pub_spec.name, type_of_pub.name" +
                " FROM type_of_pub INNER JOIN pub_spec ON type_of_pub.id_type_of_pub = pub_spec.id_type_of_pub;" },
            { "type_of_pub", "SELECT type_of_pub.name" +
                " FROM type_of_pub;" },
            { "discipline", "SELECT discipline.name" +
                " FROM discipline;" },
            { "author", "SELECT author.last_n, author.name, author.mid_n" +
                " FROM author;" },
            { "pstions", "SELECT pstions.name" +
                " FROM pstions;" },
        };

        string[] tabNames = new string[10] { "pub","give_pub","return_pub","rcpnt","num_of_pub","pub_sec","type_of_pub","discipline","author","pstions" };



        public main()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        private void main_Load(object sender, EventArgs e)
        {
            Bitmap bitmap1 = Bitmap.FromHicon(SystemIcons.Question.Handle);
            toolStripButton2.Image = bitmap1;

            conn = new OleDbConnection(connectionString);
            myDataAdapter = new OleDbDataAdapter();

            

            ViewTab("pub");

            
        }


        void ViewTab(string who)
        {


            conn.Open();
            myOleDbCommand = new OleDbCommand(selecQuerys[who], conn);
            myDataAdapter.SelectCommand = myOleDbCommand;
            myDataSet = new DataSet();
            myDataAdapter.Fill(myDataSet);
            dataGridView.DataSource = myDataSet.Tables[0];
            myDataAdapter.Update(myDataSet);

            dgvnav.BindingSource = new BindingSource(myDataSet, myDataSet.Tables[0].ToString());

            conn.Close();
        }

        private void tabs_combo_SelectedIndexChanged(object sender, EventArgs e)
        {
            ViewTab(tabNames[tabs_combo.SelectedIndex]);
        }
    }
}
