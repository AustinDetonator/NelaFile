using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace NelaFile
{
    public partial class frmElektraData : Form
    {
        //PUBLIC
        public string Order_No { get; set; }
        public string Release_No { get; set; }
        public string Sequence_No { get; set; }
        // private
        private string _C11_PACKAGE_NUM1 = String.Empty;
        private string _C13_PACKAGE_NUM2 = String.Empty;
        public frmElektraData()
        {
            InitializeComponent();
        }
        private void frmElektraData_Load(object sender, EventArgs e)
        {
            Napln_dataGridViewElektra();
        }
        private void Napln_dataGridViewElektra()
        {
            this.Cursor = Cursors.WaitCursor;

            string sCnString = Globals.ConStrin_AD9PROD__Elektra;
            DataSet ds = new DataSet();
            OracleDataAdapter ad = new OracleDataAdapter();

            //string sSQL = "select * from IFSAPP.A1135_ORDER_DETAIL a where a.CUSTOM_ORDER_LINE LIKE '" + this.Order_No + "%' order by CUSTOM_ORDER_LINE";
            string sSQL = "select * FROM TABLE(IFSAPP.A1135_ELEKTRA_API.get_Order_Detail('"+ this.Order_No + "','" + this.Release_No + "','"+ this.Sequence_No + "'))";
            OracleConnection cn = new OracleConnection(sCnString);
            cn.Open();
            OracleCommand command = new OracleCommand(sSQL);
            command.Connection = cn;
            ad.SelectCommand = command;
            ad.Fill(ds, "ORDER_DETAIL"); //C_BARCODE_SHPORD_TAB
            DataTable dt = ds.Tables[0];
            dataGridViewElektra.DataSource = dt;
            cn.Close();
            this.Cursor = Cursors.Default;
        }
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridViewElektra_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewElektra.SelectedRows.Count > 0)
            {
                DataGridViewRow currentRow = dataGridViewElektra.SelectedRows[0];
                _C11_PACKAGE_NUM1 = Convert.ToString(currentRow.Cells[11].Value);
                _C13_PACKAGE_NUM2 = Convert.ToString(currentRow.Cells[13].Value);
                textBox_C11_PACKAGE_NUM1.Text = _C11_PACKAGE_NUM1;
                textBox_C13_PACKAGE_NUM2.Text = _C13_PACKAGE_NUM2;
            }


        }

        private void button_CreateNewOrder_Click(object sender, EventArgs e)
        {
            int iC11 = Int32.Parse(textBox_C11_PACKAGE_NUM1.Text);
            int iC13 = Int32.Parse(textBox_C13_PACKAGE_NUM2.Text);

            if (iC11 < Int32.Parse(_C11_PACKAGE_NUM1)) 
            {
                MessageBox.Show("Karton č. 1 nemůže být menší než nejnižší číslo kartonu v daném řádku Zákaznické objednávky");
                textBox_C11_PACKAGE_NUM1.Text = _C11_PACKAGE_NUM1;
            }
            if (iC13 > Int32.Parse(_C13_PACKAGE_NUM2))
            {
                MessageBox.Show("Karton č. 2 nemůže být větší než nejvyšší číslo kartonu v daném řádku Zákaznické objednávky");
                textBox_C13_PACKAGE_NUM2.Text = _C13_PACKAGE_NUM2;
            }

        }
    }
}
