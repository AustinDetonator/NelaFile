using System;
using System.Net;
using System.Diagnostics;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
//using OpenNETCF.Desktop.Communication;
using System.Threading;
using System.Data.OracleClient;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

using Oracle.ManagedDataAccess.Client;

namespace NelaFile
{

    public partial class Form1 : Form
    {
        //private string Globals.TNSNAME = "AD9UPG3";
        private DataSet _DS;
        private string _ZO_BatchNo = string.Empty;
        private string _VZ_BatchNo = string.Empty;
        private string _VZ_Order_No = string.Empty;
        private string _VZ_Release_No = string.Empty;
        private string _VZ_sequence_No = string.Empty;
        private string _Customer_order = string.Empty;
        private string _LINE_NO = string.Empty;
        private string _RELEASE_NO = string.Empty;
        private string _LINE_ITEM_NO = string.Empty;
        private string _PART_NO = string.Empty;
        private string _NELA = string.Empty;
        private string _DOC_REFERENCE_OBJECT_KEY_VALUE = string.Empty;
        private int _Current_Row = -1;
        private string _FILE_NAME = string.Empty;
        //private string _FOLDER_NALEPKY = "\\\\Bap01\\ifs_nalepky\\";
        private string _FOLDER_NALEPKY = "\\\\Bap01\\ifsftp\\NALEPKY\\";
        private string _COLUMN_NAME = string.Empty;
        private int _COLUMN_INDEX = 0;
        private string _UserName = string.Empty;
        private string _ComputerName = string.Empty;
        private string _Version = string.Empty;
        private string _IPAddress = string.Empty;

        private bool _Preprog = false;

        private DataSet _dsNalepky;
        private DataTable _dtNalepky;

        private List<string> _ColumnListLabels = new List<string>();
        private List<string> _Labels_BTW_names = new List<string>();

        DataColumn[] _keysNalepky = new DataColumn[5];

        // pro alternativní nálepky (např. pro Ukrajinu), kde čk je generován z objednávky zákazníka a ne z naší
        private string _AlterCK_Order_no = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //
            _UserName = System.Windows.Forms.SystemInformation.UserName.ToString();
            tsslUsername.Text = _UserName;
            tsslComputer.Text = System.Net.Dns.GetHostName();
            tsslVersion.Text = this.GetVersion();
            tsslIPAddress.Text = this.GetIpAddress();
            tsslDB.Text = Globals.TNSNAME;
            //
            radioButton1.Checked = true;
            Globals.IFS_VERSION = IFS_Versions.IFS_9;
            //
#if DEBUG
            string applicationName = Environment.GetCommandLineArgs()[0];
#else
            string applicationName = Environment.GetCommandLineArgs()[0]; //+ ".exe";
#endif
            //
            _ComputerName = System.Net.Dns.GetHostName();
            _Version = this.GetVersion();
            _IPAddress = this.GetIpAddress();            //
            string exePath = System.IO.Path.Combine(Environment.CurrentDirectory, applicationName);
            System.Configuration.Configuration configNew = ConfigurationManager.OpenExeConfiguration(exePath);
            if (configNew.AppSettings.Settings.Count == 0)
            {
                configNew.AppSettings.Settings.Add("OutputFolder", "");
                configNew.Save(ConfigurationSaveMode.Full);

            }

            //textBoxFolderToSave.Text = configNew.AppSettings.Settings["OutputFolder"].Value;

            textBoxFolderToSave.Text = ConfigurationManager.AppSettings["OutputFolder"].ToString();  //ReadSetting("OutputFolder");
            Globals.FOLDER_BTW_EXPORT = ConfigurationManager.AppSettings["OutputFolder"].ToString() + @"\export";

            Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
            NaplnComboNelaCislo();
            Columns_NaplnList();
            //
            ProgramyAD_Log_Zapis();
        }
        public string GetIpAddress()
        {
            IPHostEntry host;
            IPAddress ip = null;

            String hostName = Dns.GetHostName();
            host = Dns.GetHostEntry(hostName);

            for (int index = 0; index < host.AddressList.Length; index++)
            {
                ip = host.AddressList[index];
            }
            return ip.ToString();
        }
        private string GetVersion()
        {
            FileVersionInfo verInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            return verInfo.FileVersion.ToString();
        }

        private void Napln_dataGridViewVZ(string pZO)
        {

            this.Cursor = Cursors.WaitCursor;

            DateTime dtB = DateTime.Now;

            //string sWhere = " to_number(to_char(a.rowversion,'YYYY')) >= 2016 and a.barcode_processed = 'FALSE' and a.customer_order like 'V%'";
            string sWhere = string.Empty;

            if (pZO != string.Empty)
            {
                sWhere = sWhere + " AND a.customer_order LIKE '" + pZO + "%'";

            }
            else
            { 
                if (checkBoxAllVZ.Checked)
                {
                    //sWhere = " AND  to_number(to_char(a.rowversion,'YYYY')) >= 2016 and a.customer_order like 'V%'";
                    //sWhere = string.Empty;

                    if (checkBoxDAOnly.Checked)
                    {
                        sWhere = sWhere + " AND substr(a.part_no,1,2) = 'DA'";

                    }
                }
                else
                {
                    if (checkBoxDAOnly.Checked)
                    {
                        sWhere = " AND substr(a.part_no,1,2) = 'DA'"; // + " AND to_number(to_char(a.rowversion,'YYYY')) >= 2016 AND a.customer_order like 'V%'";

                    }
                    else
                    { 
                        sWhere = sWhere + " AND substr(a.part_no,1,2) = 'NA'" +
                                  (sWhere==string.Empty?string.Empty: " AND " + sWhere);
                    }
                }
            }
            string sSQL = "select p.order_no || '-' || p.release_no || '-' || p.sequence_no as VZ_BatchNo," +
            "p.customer_order || '-' || p.customer_line_no || '-' || p.customer_rel_no as ZO_BatchNo," +
            "p.order_no,p.release_no,p.sequence_no," +
            "p.part_no," +
            "(SELECT ip.description FROM IFSAPP.inventory_part_TAB ip WHERE ip.part_no = p.part_no AND ip.contract = 'AD') as part_desc," +
            "p.customer_order,p.customer_line_no,p.customer_rel_no,p.customer_line_item_no,p.CelkemKs,p.customer_no," +
            "(SELECT ci.name FROM IFSAPP.customer_info_TAB ci WHERE ci.customer_id = p.customer_no) as Zakaznik,p.rowversion" +
            " FROM (" +
            "select x.order_no, x.release_no, x.sequence_no," +
            "MAX(x.customer_order) as customer_order, MAX(x.customer_line_no) as customer_line_no, MAX(x.customer_rel_no) as customer_rel_no," +
            "MAX(x.customer_line_item_no) as customer_line_item_no," +
            "MAX(x.customer_no) as customer_no," +
            "MAX(x.part_no) as part_no," +
            "SUM(x.sub_pack_amount) as CelkemKs," +
            "MAX(CASE WHEN x.unique_pkg = 'TRUE' THEN 'FALSE' ELSE 'TRUE' END) UsporneB," +
            "max(x.rowversion) as rowversion" +
            " FROM" +
            "(" +
            "  SELECT" +
            "  a.order_no, a.release_no, a.sequence_no," +
            "  a.customer_order, a.customer_line_no, a.customer_rel_no, a.customer_line_item_no, a.part_no, a.sub_pack_amount, a.customer_no" +
            "  , c.unique_pkg, a.rowversion" +
            "  FROM IFSAPP.C_BARCODE_SHPORD_TAB a," +
            "  IFSAPP.SHOP_ORD_TAB b" +
            "  , IFSAPP.C_PACKAGE_HEAD_TAB c" +
            "  WHERE 1 = 1" +
            "  AND a.order_no = b.order_no" +
            "  AND a.release_no = b.release_no" +
            "  AND a.sequence_no = b.sequence_no" +
            "  AND a.customer_order = c.order_no" +
            "  AND a.customer_line_no = c.line_no" +
            "  AND a.customer_rel_no = c.rel_no" +
            "  AND  a.rowversion >= '1.1.2019'" + sWhere + 
            ") x" +
            " GROUP BY x.order_no, x.release_no, x.sequence_no" +
            ") p" +
            " ORDER BY p.rowversion desc"; 

            //string sSQL = "SELECT x.VZ_BatchNo,x.ZO_BatchNo,x.Order_no,x.release_no,x.sequence_no,x.part_no," +
            //    "MAX((SELECT distinct ip.description FROM IFSAPP.inventory_part ip WHERE ip.part_no = x.part_no AND ip.contract = 'AD')) as part_desc," +
            //"MAX(x.customer_order) as customer_order," +
            //"MAX(x.customer_line_no) as customer_line_no," +
            //"MAX(x.customer_rel_no) as customer_rel_no," +
            //"MAX(x.customer_line_item_no) as customer_line_item_no," +
            //"SUM(x.sub_pack_amount) as CelkemKs," +
            //"MAX((SELECT ci.name FROM IFSAPP.customer_info ci WHERE ci.customer_id = x.customer_no)) as Zakaznik," +
            //"MAX(CASE WHEN x.unique_pkg = 'TRUE' THEN 'FALSE' ELSE 'TRUE' END) UsporneB" +
            //" FROM" +
            //" (" +
            //"SELECT a.order_no || '-' || a.release_no ||'-' || a.sequence_no as VZ_BatchNo," +
            //"a.customer_order || '-' || a.customer_line_no ||'-' || a.customer_rel_no as ZO_BatchNo," +
            //"a.order_no,a.release_no,a.sequence_no," +
            //"a.customer_order,a.customer_line_no,a.customer_rel_no,a.customer_line_item_no," +
            //"a.part_no," +
            //"a.sub_pack_amount," +
            //"a.customer_no,b.unique_pkg" +
            ////" FROM IFSAPP.C_BARCODE_SHPORD_TAB a,IFSAPP.C_PACKAGE_HEAD b " +
            //" FROM IFSAPP.C_BARCODE_SHPORD_TAB a,IFSAPP.C_PACKAGE_HEAD b, IFSAPP.SHOP_ORD c" +
            //" WHERE  a.customer_order = b.order_no" +
            //" AND a.customer_line_no = b.line_no" +
            //" AND a.customer_rel_no = b.rel_no" +
            //" AND a.order_no = c.order_no AND a.release_no = c.release_no AND a.sequence_no = c.sequence_no" +
            //" AND extract(year from a.rowversion) >= extract(year from sysdate) - 1 " +
            //(checkBoxAllVZ.Checked ? string.Empty : " AND c.state != 'Uzavřeno'") +
            //sWhere +
            //") x" +
            //" GROUP by x.VZ_BatchNo,x.ZO_BatchNo,x.Order_no,x.release_no,x.sequence_no,x.part_no" +
            ////" HAVING  to_number(substr(substr(x.ZO_BatchNo,2,instr(x.ZO_BatchNo,'-',1,1)-2),1,2)) > 11 " +
            //" ORDER BY to_number(substr(x.ZO_BatchNo,2,instr(x.ZO_BatchNo,'-',1,1)-2)) DESC";

            if (checkBoxOnlyStockOrders.Checked)
            {
                sSQL = "select  a.order_no || '-' || a.release_no ||'-' || a.sequence_no as VZ_BatchNo," +
                    " ' ' AS ZO_BatchNo, a.order_no,a.release_no,a.sequence_no," +
                    "a.part_no,IFSAPP.Inventory_Part_API.Get_Description(a.contract, a.Part_No) AS Part_desc," +
                    "        ' ' as customer_order," +
                    "        ' ' as customer_line_no," +
                    "        ' ' as customer_rel_no, " +
                    "        0 as customer_line_item_no," +
                    "        a.REVISED_QTY_DUE as CelkemKs," +
                    "        'Skladová objednávka' as Zakaznik, '' as UsporneB" +
                    " FROM IFSAPP.SHOP_ORD a " +
                    "         WHERE a.order_no NOT IN (select distinct b.order_no  FROM IFSAPP.CUSTOMER_ORDER b WHERE a.order_no = b.order_no)" +
                    "         AND substr(a.order_no,0,2) IN  ('P2','P3')" +
                    "      AND substr(a.part_no,1,2) = 'NA'";
            }


            //sSQL = "select part_no from IFSAPP.inventory_part_TAB WHERE part_no LIKE 'NA%'";
            //string sConString = "Data Source=AD9UPG3;User Id=IFSAPP;Password=mmLPda300m;";

            Oracle.ManagedDataAccess.Client.OracleDataAdapter ad = new Oracle.ManagedDataAccess.Client.OracleDataAdapter();

            _DS = new DataSet();
            //string sCnString = sConString; //this.ConnString;
            if (dataGridViewVZ.DataSource == null)
            {
            }
            else
            {
                DataTable dt = (DataTable)dataGridViewVZ.DataSource;
                dt.Rows.Clear();
                dataGridViewVZ.Refresh();
            }
            try
            {
                string sCnString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;";
                Oracle.ManagedDataAccess.Client.OracleConnection cn = new Oracle.ManagedDataAccess.Client.OracleConnection(sCnString);
                cn.Open();
                Oracle.ManagedDataAccess.Client.OracleCommand command = new Oracle.ManagedDataAccess.Client.OracleCommand(sSQL);
                command.Connection = cn;
                ad.SelectCommand = command;
                ad.Fill(_DS, "aa"); //C_BARCODE_SHPORD_TAB
                DataTable dt = _DS.Tables[0];
                dataGridViewVZ.DataSource = null;
                dataGridViewVZ.DataSource = dt;
                cn.Close();

                // Set the column header names.
                dataGridViewVZ.Columns[0].HeaderText = "VZ-BatchNo";
                dataGridViewVZ.Columns[0].Width = 100;
                //
                dataGridViewVZ.Columns[1].HeaderText = "ZO-BatchNo";
                dataGridViewVZ.Columns[1].Width = 80;
                //
                dataGridViewVZ.Columns[2].HeaderText = "VZ-ORDER_NO";
                dataGridViewVZ.Columns[2].Width = 100;
                //
                dataGridViewVZ.Columns[3].HeaderText = "VZ RELEASE_NO";
                dataGridViewVZ.Columns[3].Width = 90;
                //
                dataGridViewVZ.Columns[4].HeaderText = "VZ SEQUENCE_NO";
                dataGridViewVZ.Columns[4].Width = 60;
                ////
                dataGridViewVZ.Columns[5].HeaderText = "Položka";
                dataGridViewVZ.Columns[5].Width = 180;

                dataGridViewVZ.Columns[6].HeaderText = "Položka  popis";
                dataGridViewVZ.Columns[6].Width = 220;

                dataGridViewVZ.Columns[7].HeaderText = "Customer order";
                dataGridViewVZ.Columns[7].Width = 100;

                dataGridViewVZ.Columns[8].HeaderText = "Customer line no";
                dataGridViewVZ.Columns[8].Width = 100;


                dataGridViewVZ.Columns[9].HeaderText = "Customer rel no";
                dataGridViewVZ.Columns[9].Width = 100;


                dataGridViewVZ.Columns[10].HeaderText = "Customer line item no";
                dataGridViewVZ.Columns[10].Width = 100;

                dataGridViewVZ.Columns[11].HeaderText = "Celkem ks";
                dataGridViewVZ.Columns[11].Width = 60;

                dataGridViewVZ.Columns[12].HeaderText = "Zákazník";
                dataGridViewVZ.Columns[12].Width = 160;

                dataGridViewVZ.Columns[13].HeaderText = "Úsporné b.";
                dataGridViewVZ.Columns[13].Width = 80;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            DateTime dt2 = DateTime.Now;
            TimeSpan ts = (dt2 - dtB);

            string sRozdil = ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString();
            tsslDuration.Text = " čas: " + sRozdil; 

            this.Cursor = Cursors.Default;
        }

        private void Napln_dataGridViewSoubory(bool pJenVlastni)
        {
            //  uz udelane XML pro pY.listy
            string cnsSQL01 = "Persist Security Info=True;User ID=epo;Password=123456;Initial Catalog=epo;Data Source=SQL01";
            SqlConnection cnSQL01;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            cnSQL01 = new SqlConnection(cnsSQL01);
            cnSQL01.Open();
            string sSQLp = "SELECT ID,DatumExportu,Nela,Uzivatel,rtrim(ZO_BatchNo) as ZO_BatchNo, rtrim(VZ_BatchNo) as VZ_BatchNo,Cesta,DatumNaImtFtp" +
                " FROM nela" +
                (pJenVlastni ? " WHERE Uzivatel = '" + _UserName + "'" : "") +
                " ORDER BY DatumExportu desc";
            SqlCommand cmd = new SqlCommand(sSQLp, cnSQL01);
            // 
            adapter.SelectCommand = cmd;
            adapter.Fill(ds, "Nela");
            DataTable dt = ds.Tables[0];
            // dgZakazky.DisplayMember = dt.Columns("doc_no").ToString
            // dgZakazky.BindingContext = New BindingContext
            dataGridViewSoubory.DataSource = dt;
            // CheckedListBox1.DisplayMember = "doc_no"



            cnSQL01.Close();

            dataGridViewSoubory.Columns["ID"].HeaderText = "ID";
            dataGridViewSoubory.Columns["ID"].Visible = false;

            dataGridViewSoubory.Columns["DatumExportu"].HeaderText = "Datum exportu";
            dataGridViewSoubory.Columns["DatumExportu"].Width = 100;

            dataGridViewSoubory.Columns["Nela"].HeaderText = "NELA";
            dataGridViewSoubory.Columns["Nela"].Width = 70;

            dataGridViewSoubory.Columns["Uzivatel"].HeaderText = "Uživatel";
            dataGridViewSoubory.Columns["Uzivatel"].Width = 110;

            dataGridViewSoubory.Columns["ZO_BatchNo"].HeaderText = "ZO_BatchNo";
            dataGridViewSoubory.Columns["ZO_BatchNo"].Width = 120;

            dataGridViewSoubory.Columns["VZ_BatchNo"].HeaderText = "VZ_BatchNo";
            dataGridViewSoubory.Columns["VZ_BatchNo"].Width = 120;

            dataGridViewSoubory.Columns["Cesta"].HeaderText = "Cesta+soubor";
            dataGridViewSoubory.Columns["Cesta"].Width = 400;

            dataGridViewSoubory.Columns["DatumNaImtFtp"].HeaderText = "Zkopírováno na Ftp Imt";
            dataGridViewSoubory.Columns["DatumNaImtFtp"].Width = 120;




        }
        private void buttonNaplnVZ_Click(object sender, EventArgs e)
        {
            Napln_dataGridViewVZ(string.Empty);
        }

        private void buttonVytvorSoubor_Click(object sender, EventArgs e)
        {
            if (dataGridViewVZ.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte řádek");
                return;
            }
            if (comboBoxNelaCislo.SelectedItem == null)
            {
                MessageBox.Show("Není zadána NELA nebo stroj.");
                comboBoxNelaCislo.Focus();
                return;
            }
            if (comboBoxNelaCislo.SelectedItem.ToString() == string.Empty)
            {
                MessageBox.Show("Není zadána NELA nebo stroj.");
                comboBoxNelaCislo.Focus();
                return;
            }
            _NELA = comboBoxNelaCislo.SelectedItem.ToString().Trim();
            DataGridViewRow currentRow = dataGridViewVZ.SelectedRows[0];
            _VZ_BatchNo = Convert.ToString(currentRow.Cells[0].Value);
            _ZO_BatchNo = Convert.ToString(currentRow.Cells[1].Value);
            _Customer_order = Convert.ToString(currentRow.Cells[7].Value);
            _LINE_NO = Convert.ToString(currentRow.Cells[8].Value);
            _RELEASE_NO = Convert.ToString(currentRow.Cells[9].Value);
            _LINE_ITEM_NO = Convert.ToString(currentRow.Cells[10].Value);
            _PART_NO = Convert.ToString(currentRow.Cells[5].Value);

            comboBoxNelaCislo.SelectedIndex = 0;

            if (_NELA == NelaFile_MachineName.E3PR || _NELA == NelaFile_MachineName.DA001)
            {
                _Preprog = checkBoxPreprog.Checked;
                //if (_Preprog)
                //{
                //    // Najdi soubor z disku
                //    if (checkBoxProgDb.Checked)
                //    {

                //    }
                //    else
                //    { 
                //        OpenFileDialog openFileDialog1 = new OpenFileDialog
                //        {
                //            //\\austin.cz\cz\AD\SHARE\UO_business\Estar\
                //            InitialDirectory = (_NELA == NelaFile_MachineName.E3PR?Globals.FOLDER_E3PR_DEFAULT:Globals.FOLDER_DA001_DEFAULT),
                //            Title = "Browse Preprog TXT Files",

                //            CheckFileExists = true,
                //            CheckPathExists = true,

                //            DefaultExt = "txt",
                //            Filter = "txt files (*.txt)|*.txt",
                //            FilterIndex = 2,
                //            RestoreDirectory = true,

                //            ReadOnlyChecked = true,
                //            ShowReadOnly = true
                //        };

                //        if (openFileDialog1.ShowDialog() == DialogResult.OK)
                //        {
                //            if (_NELA == NelaFile_MachineName.E3PR)
                //            {
                //                Globals.FILE_E3PR_TXT = openFileDialog1.FileName;
                //            }
                //            else
                //            {
                //                Globals.FILE_DA001_TXT = openFileDialog1.FileName;
                //            }
                //        }
                //        else
                //        {
                //            return;
                //        }
                //    }
                //}
                VytvorSouborProDA();
            }
            else
            {
                VytvorSoubor();
            }

        }
        private void VytvorSoubor()
        {
            this.Cursor = Cursors.WaitCursor;

            string sConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;"; 
            //OracleConnection cn;
            //OracleCommand cm;
            //OracleDataReader rd;
            StreamWriter sw;

            string sLine;
            string sBatchNo = _ZO_BatchNo;
            string sPACKAGE_QTY_MULTI = string.Empty;
            string sNOMINAL_TIME_NA_MULTI = string.Empty;
            string sDelay_step_multi = string.Empty;
            string sAUTO_LABEL = string.Empty;
            string sCatalog_no = string.Empty;
            string sCatalog_desc = string.Empty;

            bool bAlter_CK = false;
            string sFile = (textBoxFolderToSave.Text.Trim() + ("\\NELA_"
            + (_VZ_BatchNo + ".TXT")));

            sw = new StreamWriter(sFile, false, System.Text.Encoding.Default);


            try
            {
                int i = 0;
                // 
                System.Data.OracleClient.OracleConnection cn = new System.Data.OracleClient.OracleConnection(sConString);
                cn.Open();

                // jedná se o úsporné balení ?
                Funkce ff = new Funkce();
                string sErr = string.Empty;
                string sSQL_UsporneBaleni = "SELECT a.order_no FROM  IFSAPP.C_PACKAGE_HEAD a" +
                            " WHERE a.unique_pkg = 'FALSE'" +
                            " AND  a.order_no = '" + _Customer_order + "'" +
                            " AND  a.line_no = '" + _LINE_NO + "'" +
                            " AND  a.rel_no = '" + _RELEASE_NO + "'";

                bool bUsporneBaleni = ff.Oracle_Exists(sConString, sSQL_UsporneBaleni, ref sErr);
                if (bUsporneBaleni)
                {
                    string sSQL_Multi = "SELECT * FROM (SELECT distinct order_no,package_num2,part_no," +
                        "replace(wm_concat(amount) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') PACKAGE_QTY_MULTI," +
                        "replace(wm_concat(nominal_time) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') NOMINAL_TIME_NA_MULTI," +
                        "replace(wm_concat(delay_step_multi) OVER (PARTITION BY package_num2 ORDER BY order_no,package_num2) ,',','+') Delay_step_multi" +
                        " FROM" +
                        "( " +
                        "   SELECT a.order_no,b.package_num2,b.ch_line_item_no,b.amount,b.part_no," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'N016') as nominal_time," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'L004') as delay_step_multi" +
                        "   FROM  IFSAPP.C_PACKAGE_HEAD a, IFSAPP.C_PACKAGE_DETAIL b " +
                        "   WHERE a.order_no = b.order_no" +
                        "   AND a.line_no = b.line_no" +
                        "   AND a.rel_no = b.rel_no" +
                        "   AND a.line_item_no = b.line_item_no" +
                        "   AND a.unique_pkg = 'FALSE'" +
                        "   AND  a.order_no = '" + _Customer_order + "'" +
                        "   AND  a.line_no = '" + _LINE_NO + "'" +
                        "   AND  a.rel_no = '" + _RELEASE_NO + "'" +
                        "   AND  a.line_item_no = '" + _LINE_ITEM_NO + "'" +
                        "   ORDER BY order_no,ch_line_item_no" +
                        "))" +
                        " WHERE part_no = '" + _PART_NO + "'" +
                        " ORDER BY order_no,package_num2";
                    // při úsporném balení vrátí více řádků (kartonů) pro výrobu jedné položky 
                    sPACKAGE_QTY_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "PACKAGE_QTY_MULTI", ref sErr);
                    sNOMINAL_TIME_NA_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "NOMINAL_TIME_NA_MULTI", ref sErr);
                    sDelay_step_multi = ff.Oracle_GetValue(sConString, sSQL_Multi, "delay_step_multi", ref sErr);

                }
                // 
                // zkontroluj country
                string sSQL_Country = "select a.country from IFSAPP.A0000_CustomerByOrderNo a where a.ORDER_NO = '" + _Customer_order + "'";
                string sCountry = ff.Oracle_GetValue(sConString, sSQL_Country, ref sErr);
                if (sErr != string.Empty)
                {
                    MessageBox.Show("Nenašel Country v zák. objednávce", sErr);
                    this.Cursor = Cursors.Default;
                    return;
                }
                if (sCountry == "UA")
                {
                    DialogResult ss= MessageBox.Show("Chcete změnit soubor podle Ukrajiny?","UKRAJINA" , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (ss == DialogResult.Yes)
                    {
                        checkBox_AlterCK.Checked = true;
                    }
                    else
                    {
                        checkBox_AlterCK.Checked = false;
                    }
                }


                if (checkBox_AlterCK.Checked)
                {
                    //if (sCountry == "UA")
                    //{
                        string sSQL_Get__AlterCK_Order_no = "select REGEXP_SUBSTR(a.alt_package_barcode_id, '[^-]+', 1, 1) ||'-' ||REGEXP_SUBSTR(a.alt_package_barcode_id, '[^-]+', 1, 2)" +
                            " from IFSAPP.c_barcode_shpord_tab a where rownum = 1 and a.package_barcode_id LIKE '" + _Customer_order + "%'";
                        _AlterCK_Order_no = ff.Oracle_GetValue(sConString, sSQL_Get__AlterCK_Order_no, ref sErr);
                        if (sErr != string.Empty)
                        {
                            MessageBox.Show(sErr);
                            this.Cursor = Cursors.Default;
                            return;
                        }
                        if (_AlterCK_Order_no=="-" || _AlterCK_Order_no == string.Empty)
                        {
                            MessageBox.Show("Naní zadán alternativní čárový kód!" + Environment.NewLine + "Zkontrolujte zák. objednávku " + _Customer_order +
                                ", zda má připojené XML a proběhla generace nálepek." + Environment.NewLine + Environment.NewLine + 
                                "Nebo odtrhněte požití alternativního čk."
                                , "Info",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                            this.Cursor = Cursors.Default;
                            return;
                        }   
                        bAlter_CK = true;
                        sBatchNo = _AlterCK_Order_no + "-1";
                        //checkBox_AlterCK.Checked = false;
                    //
                    // checkBox_AlterCK.Checked = false;
                }
                //
                sLine = "[TABLE]";
                sw.WriteLine(sLine);
                sLine = "MACHINE_NAME = " + _NELA;
                sw.WriteLine(sLine);
                sLine = "STATION_NUM = " + _NELA;
                sw.WriteLine(sLine);
                sLine = ("BATCH_NO = " + sBatchNo);
                sw.WriteLine(sLine);
                // KARTONY
                // Dim sSQL_K As String = "SELECT batch_no,doc_class,MAX(file_name) AS FileName  FROM IFSAPP.c_labels_ad a" & _
                //       " WHERE batch_no = '" & sBatchNo & "'  AND doc_class = 'N-NA-KTN'" & _
                //       " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"
                string sSQL_K = ("SELECT distinct file_name FROM IFSAPP.c_labels_ad a" + (" WHERE a.order_no = \'"
                            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%KTN%'" +
                            " AND (upper(file_name) like '%.BTW' OR upper(file_name) like '%.ZPL')" +
                            " ORDER BY a.file_name"))))))))))))));
                System.Data.OracleClient.OracleCommand cmdK = new System.Data.OracleClient.OracleCommand(sSQL_K, cn);
                System.Data.OracleClient.OracleDataReader rdK = cmdK.ExecuteReader();
                if (!rdK.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-KTN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                        sw.WriteLine(sLine);
                    }
                }
                while (rdK.Read())
                {
                    i++;
                    if ((i == 2))
                    {
                        if (rdK[0].GetType() != typeof(DBNull))
                        {
                            sLine = ("FILENAME_KTN2 = " + rdK.GetString(0).Trim());
                        }
                        else
                        {
                            sLine = "FILENAME_KTN2 = ";
                        }
                    }
                    else if ((rdK[0].GetType() != typeof(DBNull)))
                    {
                        sLine = ("FILENAME_KTN = " + rdK.GetString(0).Trim());
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                    }
                    sw.WriteLine(sLine);
                }
                if ((i == 1))
                {
                    sLine = "FILENAME_KTN2 =";
                    sw.WriteLine(sLine);
                }
                rdK.Close();
                cmdK.Dispose();


                //  pro Skladacku
                string sSQL_S = string.Empty;
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_S = Globals.SQL_Nalepky_From_DMS_IFS9 +
                        _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%STN%'" +
                        " AND (upper(d.file_name) like '%.BTW' OR upper(d.file_name) like '%.ZPL')";

                        ;
                }
                else
                { 
                     sSQL_S = ("SELECT MAX(file_name) AS FileName,batch_no,doc_class FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'" 
                            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%STN%'" 
                            + " AND (upper(file_name) like '%.BTW' OR upper(file_name) like '%.ZPL')" 
                            + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));
                }
                System.Data.OracleClient.OracleCommand cmdS = new System.Data.OracleClient.OracleCommand(sSQL_S, cn);
                System.Data.OracleClient.OracleDataReader rdS = cmdS.ExecuteReader();
                if (!rdS.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-STN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_STN = ";
                        sw.WriteLine(sLine);
                    }
                }
                sLine = "FILENAME_STN = ";
                while (rdS.Read())
                {
                    if (rdS[0].GetType() == typeof(DBNull))
                    {
                        sLine = "FILENAME_STN = ";
                    }
                    else
                    {
                        sLine = ("FILENAME_STN = " + rdS.GetString(0).Trim());
                        //AUTO_LABEL
                        string sHodnotaPole = rdS.GetString(0).Trim();
                        //-A1 - 1.ZPL
                        sHodnotaPole = sHodnotaPole.Replace("-A1-1.BTW", "");
                        sHodnotaPole = sHodnotaPole.Replace("-A1.BTW", "");
                        string sSQL_AUTO_LABEL = "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'";
                        sAUTO_LABEL = MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, sSQL_AUTO_LABEL);
                        sAUTO_LABEL = "AUTO_LABEL = " + sAUTO_LABEL;
                    }
                    sw.WriteLine(sLine);
                }
                rdS.Close();
                cmdS.Dispose();

                //  pro Praporek

                string sSQL_x = ("SELECT batch_no,doc_class,MAX(file_name) AS FileName FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%PRO%'"
                            + " AND (upper(file_name) like '%.BTW' OR upper(file_name) like '%.ZPL')"
                            + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));

                string sSQL_P = string.Empty;
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_P = Globals.SQL_Nalepky_From_DMS_IFS9 +
                        _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";
                }
                else
                { 
                    sSQL_P = "SELECT (c.doc_class || '-' || c.doc_no || '-' ||c.doc_rev || '.' || c.format_size) as Nalepka" +
                        " FROM IFSAPP.DOC_REFERENCE_OBJECT a, IFSAPP.DOC_STRUCTURE b,IFSAPP.DOC_ISSUE_REFERENCE c" +
                        " WHERE a.doc_class = b.doc_class" +
                        " AND a.doc_no = b.doc_no" +
                        " AND a.doc_rev = b.doc_rev" +
                        " AND b.sub_doc_class = c.doc_class" +
                        " AND b.sub_doc_no  = c.doc_no" +
                        " AND b.sub_doc_rev = c.doc_rev" +
                        " AND a.lu_name = 'CustomerOrderLine'" +
                        //" AND a.key_ref LIKE  '%ORDER_NO=" + _Customer_order + "%'" +
                        //" AND a.key_ref LIKE  '%LINE_NO=" + _LINE_NO + "%'" +
                        //" AND a.key_ref LIKE  '%REL_NO=" + _RELEASE_NO + "%'" +
                        " AND a.key_value LIKE  '" + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";

                }
                System.Data.OracleClient.OracleCommand cmdP = new System.Data.OracleClient.OracleCommand(sSQL_P, cn);
                System.Data.OracleClient.OracleDataReader rdP = cmdP.ExecuteReader();
                if (!rdP.HasRows) // NOT !!!!!!!!!!!
                {
                    if ((checkBoxOnlyStockOrders.Checked == true))
                    {
                        sLine = "FILENAME_PRO = ";
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        throw new Exception("Nenašel nálepky PRO !!");
                    }
                }
                else // má řádky
                {
                    int iRows = 0;
                    bool bPRO_BTW = false;
                    while (rdP.Read())
                    {
                        iRows++;
                    }
                    if (iRows > 1) // Zvol BTW nebo ZPL 
                    {
                        if (MessageBox.Show("Nalezeny 2 nálepky pro praporek, chcete zařadit do txt praporek BTW ?" +
                            Environment.NewLine + "Pokud ne, vezme se ZPL", "BTW nebo ZPL", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            bPRO_BTW = true;
                        }
                    }
                    rdP = cmdP.ExecuteReader();
                    i = 0;
                    while (rdP.Read())
                    {
                        i++;
                        //
                        if (rdP[0].GetType() == typeof(DBNull))
                        {
                            sLine = "FILENAME_PRO = ";
                        }
                        else
                        {
                            string sPRO = rdP.GetString(0).Trim();
                            if (iRows == 1) // existuje jen l řádek
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sw.WriteLine(sLine);
                            }
                            if (sPRO.Contains(".BTW") && bPRO_BTW == true && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO2 = " + rdP.GetString(0).Trim());
                                sw.WriteLine(sLine);
                            }
                            if (sPRO.Contains(".ZPL") && bPRO_BTW == false && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO2 = " + rdP.GetString(0).Trim());
                                sw.WriteLine(sLine);
                            }
                        }
                    }
                    rdP.Close();
                }
                cmdP.Dispose();
                // 
                // Dim sSQL As String = DejSQL_CLabels(sBatchNo)
                // 
                // Uložená procedura IFSAPP.A0055_ONLINENELA.PRONELU
                string sStoredProcedure = string.Empty;
                
                //sConString = (Globals.IFS_VERSION == IFS_Versions.IFS_9 ? Globals.ConString_IFS9 : "Data Source=ADPROD;User ID=lotyl;Password=123456");

                System.Data.OracleClient.OracleConnection cnO = new System.Data.OracleClient.OracleConnection(sConString);

                sStoredProcedure = "IFSAPP.A0055_ONLINENELA.PRONELU";

                System.Data.OracleClient.OracleCommand cmdO = new System.Data.OracleClient.OracleCommand(sStoredProcedure, cnO);
                cmdO.CommandType = CommandType.StoredProcedure;
                cmdO.CommandText = "IFSAPP.A0055_ONLINENELA.PRONELU";
                cmdO.CommandType = CommandType.StoredProcedure;
                // cmdO.Parameters.Add(New OracleParameter("PZOBATCHNO", OracleType.VarChar, 50)).Value = _ZO_BatchNo
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("PVZBATCHNO", OracleType.VarChar, 50)).Value = _VZ_BatchNo;
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("pOrder_no", OracleType.VarChar, 50)).Value = _Customer_order;
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("pLine_no", OracleType.VarChar, 50)).Value = _LINE_NO;
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("pRelease_no", OracleType.VarChar, 50)).Value = _RELEASE_NO;
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("pLine_item_no", OracleType.VarChar, 50)).Value = _LINE_ITEM_NO;
                cmdO.Parameters.Add(new System.Data.OracleClient.OracleParameter("PRESULTSET", OracleType.Cursor)).Direction = ParameterDirection.Output;
                DataSet ds = new DataSet();
                System.Data.OracleClient.OracleDataAdapter da = new System.Data.OracleClient.OracleDataAdapter();
                da.SelectCommand = cmdO;
                da.Fill(ds);
                DataTable dt = ds.Tables[0];

                //rdO = cmdO.Parameters["pResultSet"].Value;
                // 
                if (dt.Rows.Count == 0)
                {
                    throw new Exception("Nenašel  v proceduře IFSAPP.A0055_ONLINENELA.PRONELU  data " + 
                        Environment.NewLine + "VZ_BatchNo :" + _VZ_BatchNo +
                        Environment.NewLine + "ZO_BatchNo :" + _ZO_BatchNo);
                }
                foreach (DataRow row in dt.Rows)
                {
                    if ((row[1].GetType() == typeof(DBNull)))
                    {
                        sLine = "CUSTOMER_PO_NO = ";
                    }
                    else
                    {
                        sLine = ("CUSTOMER_PO_NO = " + row["CUSTOMER_PO_NO"].ToString());
                    }
                    sw.WriteLine(sLine);

                    // 
                    if (row["ORDER_NO"].GetType() == typeof(DBNull))
                    {
                        sLine = "ORDER_NO = ";
                    }
                    else
                    {
                        sLine = ("ORDER_NO = " + row["ORDER_NO"].ToString());
                        if (bAlter_CK)
                        {
                            sLine = "ORDER_NO = " + _AlterCK_Order_no.Split('-')[0].ToString();

                        }
                    }

                    sw.WriteLine(sLine);
                    // 
                    if ((row["LINE_NO"].GetType() == typeof(DBNull)))
                    {
                        sLine = "LINE_NO = ";
                    }
                    else
                    {
                        sLine = ("LINE_NO = " + row["LINE_NO"].ToString());
                        if (bAlter_CK)
                        {
                            sLine = "LINE_NO = " + _AlterCK_Order_no.Split('-')[1].ToString();
                        }

                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["RELEASE_NO"].GetType() == typeof(DBNull)))
                    {
                        sLine = "RELEASE_NO = ";
                    }
                    else
                    {
                        sLine = ("RELEASE_NO = " + row["RELEASE_NO"].ToString());
                        if (bAlter_CK)
                        {
                            sLine = "RELEASE_NO = 1";
                        }
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["LINE_ITEM_NO"].GetType() == typeof(DBNull)))
                    {
                        sLine = "LINE_ITEM_NO = ";
                    }
                    else
                    {
                        sLine = ("LINE_ITEM_NO = " + row["LINE_ITEM_NO"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_NUM1"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_NUM1 = ";
                    }
                    else
                    {
                        sLine = ("PACKAGE_NUM1 = " + row["PACKAGE_NUM1"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_NUM2"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_NUM2 = ";
                    }
                    else
                    {
                        sLine = ("PACKAGE_NUM2 = " + row["PACKAGE_NUM2"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["CATALOG_NO"].GetType() == typeof(DBNull)))
                    {
                        sLine = "CATALOG_NO = ";
                    }
                    else
                    {
                        if (checkBox_AlterCK.Checked)
                        {
                            sCatalog_no = XML_ExLabeling_Get_Part_no(_Customer_order, row["CATALOG_NO"].ToString());
                            if (sCatalog_no == string.Empty)
                            { 
                                MessageBox.Show("Děláte soubor pro Ukrajinu a nenašel jsem náhradní katalogové číslo pro " + row["CATALOG_NO"].ToString() + Environment.NewLine +
                                    "Zkontrolujte, zda je v XML připojeném k zák. objednávce " + _Customer_order + " uvedeno náhradní katalogové číslo pro tento díl.", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                return;
                            }
                            else
                            {
                                sLine = ("CATALOG_NO = " + sCatalog_no);
                            }
                        }
                        else
                        {
                            sLine = ("CATALOG_NO = " + row["CATALOG_NO"].ToString());
                        }

                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["CATALOG_DESC"].GetType() == typeof(DBNull)))
                    {
                        sLine = "CATALOG_DESC = ";
                    }
                    else
                    {
                        sCatalog_desc = row["CATALOG_DESC"].ToString();
                        sLine = ("CATALOG_DESC = " + sCatalog_desc);
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_QTY"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_QTY = ";
                    }
                    else
                    {
                        sLine = ("PACKAGE_QTY = " + row["PACKAGE_QTY"].ToString());
                    }
                    sw.WriteLine(sLine);
                    //
                    //if (bUsporneBaleni)
                    //{
                    //97_PACKAGE_QTY_MULTI;98_NOMINAL_TIME_NA_MULTI;99_DELAY_STEP_MULTI 
                    sLine = "PACKAGE_QTY_MULTI = " + sPACKAGE_QTY_MULTI;
                    sw.WriteLine(sLine);
                    //}
                    // 
                    if ((row["SMALL_PACKAGE_QTY"].GetType() == typeof(DBNull)))
                    {
                        sLine = "SMALL_PACKAGE_QTY = ";
                    }
                    else
                    {
                        sLine = ("SMALL_PACKAGE_QTY = " + row["SMALL_PACKAGE_QTY"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_GROSS_WEIGHT"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_GROSS_WEIGHT = ";
                    }
                    else
                    {
                        sLine = "PACKAGE_GROSS_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(row["PACKAGE_GROSS_WEIGHT"].ToString()));
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["ROUND_NET_WEIGHT"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ROUND_NET_WEIGHT = ";
                    }
                    else
                    {
                        sLine = ("ROUND_NET_WEIGHT = " + row["ROUND_NET_WEIGHT"].ToString());
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_NET_WEIGHT"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_NET_WEIGHT = ";
                    }
                    else
                    {
                        //sLine = ("PACKAGE_NET_WEIGHT = " + String.Format("#####0.0",row["PACKAGE_NET_WEIGHT"]));
                        sLine = "PACKAGE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:#####0.0}", Convert.ToDouble(row["PACKAGE_NET_WEIGHT"].ToString()));
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["EXPLOSIVE_NET_WEIGHT"].GetType() == typeof(DBNull)))
                    {
                        sLine = "EXPLOSIVE_NET_WEIGHT = ";
                    }
                    else
                    {
                        //sLine = ("EXPLOSIVE_NET_WEIGHT = " + String.Format( "#####0.000",row["EXPLOSIVE_NET_WEIGHT"]));
                        sLine = "EXPLOSIVE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:#####0.000}", Convert.ToDouble(row["EXPLOSIVE_NET_WEIGHT"].ToString()));
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["UN_NO"].GetType() == typeof(DBNull)))
                    {
                        sLine = "UN_NO = ";
                    }
                    else
                    {
                        sLine = ("UN_NO = " + row["UN_NO"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["ADR_RID_CLASS"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ADR_RID_CLASS = ";
                    }
                    else
                    {
                        sLine = ("ADR_RID_CLASS = " + row["ADR_RID_CLASS"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["DETONATOR_CLASS"].GetType() == typeof(DBNull)))
                    {
                        sLine = "DETONATOR_CLASS = ";
                    }
                    else
                    {
                        sLine = ("DETONATOR_CLASS = " + row["DETONATOR_CLASS"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // pro EA - nebrat
                    //if ((row["DELAY_LEVEL"].GetType() == typeof(DBNull))) {sLine = "DELAY_LEVEL = ";} else {sLine = ("DELAY_LEVEL = " + row["DELAY_LEVEL"].ToString());}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["CAVITY_MATERIAL"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "CAVITY_MATERIAL = ";
                    //}
                    //else
                    //{
                    //    sLine = ("CAVITY_MATERIAL = " + row["CAVITY_MATERIAL"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["IGNITION_CUE_MJ"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "IGNITION_CUE_MJ = ";
                    //}
                    //else
                    //{
                    //    sLine = ("IGNITION_CUE_MJ = " + row["IGNITION_CUE_MJ"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["IGNITION_CUE_J"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "IGNITION_CUE_J = ";
                    //}
                    //else
                    //{
                    //    sLine = ("IGNITION_CUE_J = " + row["IGNITION_CUE_J"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["SAFETY_CUE_MJ"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "SAFETY_CUE_MJ = ";
                    //}
                    //else
                    //{
                    //    sLine = ("SAFETY_CUE_MJ = " + row["SAFETY_CUE_MJ"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["SAFETY_CUE_J"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "SAFETY_CUE_J = ";
                    //}
                    //else
                    //{
                    //    sLine = ("SAFETY_CUE_J = " + row["SAFETY_CUE_J"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["SUPPLY_WIRE_MATERIAL"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "SUPPLY_WIRE_MATERIAL = ";
                    //}
                    //else
                    //{
                    //    sLine = ("SUPPLY_WIRE_MATERIAL = " + row["SUPPLY_WIRE_MATERIAL"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["WIRE_ISULATOR_MATERIAL"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "WIRE_ISULATOR_MATERIAL = ";
                    //}
                    //else
                    //{
                    //    sLine = ("WIRE_ISULATOR_MATERIAL = " + row["WIRE_ISULATOR_MATERIAL"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["WIRE_LENGTH"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "WIRE_LENGTH = ";
                    //}
                    //else
                    //{
                    //    sLine = ("WIRE_LENGTH = " + row["WIRE_LENGTH"].ToString());
                    //}
                    //sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    //sw.WriteLine(sLine);
                    // 
                    //if ((row["WIRE_LENGTH_TRACE"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "WIRE_LENGTH_TRACE = ";
                    //}
                    //else
                    //{
                    //    sLine = ("WIRE_LENGTH_TRACE = " + row["WIRE_LENGTH_TRACE"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    if ((row["CONSUMPTION_DATE"].GetType() == typeof(DBNull)))
                    {
                        sLine = "CONSUMPTION_DATE = ";
                    }
                    else
                    {
                        sLine = ("CONSUMPTION_DATE = " + row["CONSUMPTION_DATE"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    //if ((row["NOMINAL_TIME_EA"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "NOMINAL_TIME_EA = ";
                    //}
                    //else
                    //{
                    //    sLine = ("NOMINAL_TIME_EA = " + row["NOMINAL_TIME_EA"].ToString());
                    //}
                    //sw.WriteLine(sLine);
                    // 
                    if ((row["NOMINAL_TIME_NA"].GetType() == typeof(DBNull)))
                    {
                        sLine = "NOMINAL_TIME_NA = ";
                    }
                    else
                    {
                        sLine = ("NOMINAL_TIME_NA = " + row["NOMINAL_TIME_NA"].ToString());
                    }

                    sw.WriteLine(sLine);
                    //
                    if (bUsporneBaleni)
                    {
                        //97_PACKAGE_QTY_MULTI;98_NOMINAL_TIME_NA_MULTI;99_DELAY_STEP_MULTI 
                        sLine = "NOMINAL_TIME_NA_MULTI = " + sNOMINAL_TIME_NA_MULTI;
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        sLine = "NOMINAL_TIME_NA_MULTI = ";
                        sw.WriteLine(sLine);
                    }

                    // 
                    if ((row["ST_LENGTH"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ST_LENGTH = ";
                    }
                    else
                    {
                        sLine = ("ST_LENGTH = " + row["ST_LENGTH"].ToString());
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["DELAY_STEP"].GetType() == typeof(DBNull)))
                    {
                        sLine = "DELAY_STEP = ";
                    }
                    else
                    {
                        sLine = row["DELAY_STEP"].ToString();
                        if (sLine.StartsWith("0") && sLine.Length > 1)
                        {
                            sLine = sLine.Substring(1, sLine.Length - 1);
                        }
                        sLine = "DELAY_STEP = " + sLine;
                    }
                    sw.WriteLine(sLine);
                    //
                    //if (bUsporneBaleni)
                    //{
                    //97_PACKAGE_QTY_MULTI;98_NOMINAL_TIME_NA_MULTI;99_DELAY_STEP_MULTI 
                    sLine = "DELAY_STEP_MULTI  = " + sDelay_step_multi;
                    sw.WriteLine(sLine);
                    //}
                    if (bUsporneBaleni)
                    {
                        sLine = "SEQ_NUM_SHIFT = 2";
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        sLine = "SEQ_NUM_SHIFT =";
                        sw.WriteLine(sLine);
                    }
                    // 
                    if ((row["TESTING_STATION_NUMBER"].GetType() == typeof(DBNull)))
                    {
                        sLine = "TESTING_STATION_NUMBER = ";
                    }
                    else
                    {
                        sLine = ("TESTING_STATION_NUMBER = " + row["TESTING_STATION_NUMBER"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["TESTING_PROTOCOL_NUMBER"].GetType() == typeof(DBNull)))
                    {
                        sLine = "TESTING_PROTOCOL_NUMBER = ";
                    }
                    else
                    {
                        sLine = ("TESTING_PROTOCOL_NUMBER = " + row["TESTING_PROTOCOL_NUMBER"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["CUSTOMER_AUTHORIZE_DECREE"].GetType() == typeof(DBNull)))
                    {
                        sLine = "CUSTOMER_AUTHORIZE_DECREE = ";
                    }
                    else
                    {
                        sLine = ("CUSTOMER_AUTHORIZE_DECREE = " + row["CUSTOMER_AUTHORIZE_DECREE"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["ANOTHER_NOTICE"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ANOTHER_NOTICE = ";
                    }
                    else
                    {
                        sLine = ("ANOTHER_NOTICE = " + row["ANOTHER_NOTICE"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["NOTE"].GetType() == typeof(DBNull)))
                    {
                        sLine = "NOTE = ";
                    }
                    else
                    {
                        sLine = ("NOTE = " + row["NOTE"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["GENERAL"].GetType() == typeof(DBNull)))
                    {
                        sLine = "GENERAL = ";
                    }
                    else
                    {
                        sLine = ("GENERAL = " + row["GENERAL"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["GENERAL_1"].GetType() == typeof(DBNull)))
                    {
                        sLine = "GENERAL_1 = ";
                    }
                    else
                    {
                        sLine = ("GENERAL_1 = " + row["GENERAL_1"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // Nove
                    if ((row["ST_COLOR"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ST_COLOR = ";
                    }
                    else
                    {
                        sLine = ("ST_COLOR = " + row["ST_COLOR"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["ST_PRODUCER"].GetType() == typeof(DBNull)))
                    {
                        sLine = "ST_PRODUCER = ";
                    }
                    else
                    {
                        sLine = ("ST_PRODUCER = " + row["ST_PRODUCER"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["STOPPER"].GetType() == typeof(DBNull)))
                    {
                        sLine = "STOPPER = ";
                    }
                    else
                    {
                        sLine = ("STOPPER = " + row["STOPPER"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["CONNECTOR"].GetType() == typeof(DBNull)))
                    {
                        sLine = "CONNECTOR = ";
                    }
                    else
                    {
                        sLine = ("CONNECTOR = " + row["CONNECTOR"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["SEALING"].GetType() == typeof(DBNull)))
                    {
                        sLine = "SEALING = ";
                    }
                    else
                    {
                        sLine = ("SEALING = " + row["SEALING"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["DETONATOR_LENGTH"].GetType() == typeof(DBNull)))
                    {
                        sLine = "DETONATOR_LENGTH = ";
                    }
                    else
                    {
                        sLine = ("DETONATOR_LENGTH = " + row["DETONATOR_LENGTH"].ToString());
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sw.WriteLine(sLine);
                    // 
                    if ((row["FLAG_COLOR"].GetType() == typeof(DBNull)))
                    {
                        sLine = "FLAG_COLOR = ";
                    }
                    else
                    {
                        sLine = ("FLAG_COLOR = " + row["FLAG_COLOR"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // TN_COLOR
                    if ((row["TN_COLOR"].GetType() == typeof(DBNull)))
                    {
                        sLine = "TN_COLOR = ";
                    }
                    else
                    {
                        sLine = ("TN_COLOR = " + row["TN_COLOR"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["manufacture_date"].GetType() == typeof(DBNull)))
                    {
                        sLine = "DATE = ";
                    }
                    else
                    {
                        sLine = ("DATE = " + row["manufacture_date"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["expiration_date"].GetType() == typeof(DBNull)))
                    {
                        sLine = "EXPIRATION = ";
                    }
                    else
                    {
                        sLine = ("EXPIRATION = " + row["expiration_date"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_NUM2BEG"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_NUM2BEG = ";
                    }
                    else
                    {
                        sLine = ("PACKAGE_NUM2BEG = " + row["PACKAGE_NUM2BEG"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["PACKAGE_NUM2END"].GetType() == typeof(DBNull)))
                    {
                        sLine = "PACKAGE_NUM2END = ";
                    }
                    else
                    {
                        sLine = ("PACKAGE_NUM2END = " + row["PACKAGE_NUM2END"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["SEQ_TOTAL"].GetType() == typeof(DBNull)))
                    {
                        sLine = "SEQ_TOTAL = ";
                    }
                    else
                    {
                        sLine = ("SEQ_TOTAL = " + row["SEQ_TOTAL"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // BOX_PCS = SEQ_TOTAL
                    if ((row["SEQ_TOTAL"].GetType() == typeof(DBNull)))
                    {
                        sLine = "BOX_PCS  = ";
                    }
                    else
                    {
                        sLine = ("BOX_PCS  = " + row["SEQ_TOTAL"].ToString());
                    }
                    sw.WriteLine(sLine);
                    // 
                    if ((row["TOTAL_PCS"].GetType() == typeof(DBNull)))
                    {
                        sLine = "TOTAL_PCS = ";
                    }
                    else
                    {
                        sLine = ("TOTAL_PCS = " + row["TOTAL_PCS"].ToString());
                    }
                    sw.WriteLine(sLine);
                    sLine = "PRODUCER = CZ010";
                    sw.WriteLine(sLine);

                    // 86_DMX_PACKAGE_QTY;
                    // 87_DMX_DELAY_LEVEL = "97";
                    // 88_DMX_DELAY_STEP;
                    // 90_DMX_NOMINAL_TIME_NA;
                    // 92_DMX_ST_LENGTH;
                    // 93_DMX_DATE;
                    // 94_DMX_PACKAGE_NUM2;
                    // 95_DMX_SEQ_NUM;
                    // 96_DMX_PRODUCER 
                    sLine = ("DMX_PRODUCER = 90CZ010" + '');
                    sw.WriteLine(sLine);
                    if ((row["PACKAGE_NUM2"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["PACKAGE_NUM2"].ToString();
                    }
                    sLine = ("DMX_PACKAGE_NUM2 = "
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    if ((row["PACKAGE_QTY"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["PACKAGE_QTY"].ToString();
                    }
                    sLine = ("DMX_PACKAGE_QTY = 37"
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    //
                    //if ((row["DELAY_LEVEL"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "";
                    //}
                    //else
                    //{
                    //    sLine = row["DELAY_LEVEL"].ToString();
                    //}
                    //sLine = ("DMX_DELAY_LEVEL = 97"
                    //            + (sLine + ''));
                    //sw.WriteLine(sLine);
                    //
                    if ((row["DELAY_STEP"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["DELAY_STEP"].ToString();
                        if (sLine.StartsWith("0") && sLine.Length > 1)
                        {
                            sLine = sLine.Substring(1, sLine.Length - 1);
                        }

                    }
                    sLine = ("DMX_DELAY_STEP = 97"
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    //
                    //if ((row["NOMINAL_TIME_EA"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "";
                    //}
                    //else
                    //{
                    //    sLine = row["NOMINAL_TIME_EA"].ToString();
                    //}
                    //sLine = ("DMX_NOMINAL_TIME_EA = 98"
                    //            + (sLine + ''));
                    //sw.WriteLine(sLine);
                    //
                    if ((row["NOMINAL_TIME_NA"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["NOMINAL_TIME_NA"].ToString();
                    }
                    sLine = ("DMX_NOMINAL_TIME_NA = 98"
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    //
                    //if ((row["WIRE_LENGTH"].GetType() == typeof(DBNull)))
                    //{
                    //    sLine = "";
                    //}
                    //else
                    //{
                    //    sLine = row["WIRE_LENGTH"].ToString();
                    //}
                    //sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    //sLine = ("DMX_WIRE_LENGTH = 99"
                    //            + (sLine + ''));
                    //sw.WriteLine(sLine);
                    //
                    if ((row["ST_LENGTH"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["ST_LENGTH"].ToString();
                    }
                    sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(",", ".") : sLine);
                    sLine = ("DMX_ST_LENGTH = 99"
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    if ((row["MANUFACTURE_DATE"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["MANUFACTURE_DATE"].ToString();
                    }
                    sLine = ("DMX_DATE = "
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                    if ((row["PocetSkl"].GetType() == typeof(DBNull)))
                    {
                        sLine = "";
                    }
                    else
                    {
                        sLine = row["PocetSkl"].ToString();
                    }
                    sLine = ("DMX_SEQ_NUM = "
                                + (sLine + ''));
                    sw.WriteLine(sLine);
                }
                //AUTO_LABEL - nové aplikační identifikátory (je to v SL), podle nich si TRC vybírá praporek


                // vyřazeno 24.3.2021 pro NELY
                //sw.WriteLine(sAUTO_LABEL);

                // zavedeno 04.04.2025 pro NELY, nějaké aplikace na čtení DMX našich kartonů nesmí mít prázdný GS92
                // sw.WriteLine("AUTO_LABEL = ");
                sw.WriteLine("AUTO_LABEL = 0000");
                //

                // UFI_CODE
                string sSQL_UFI = "select ch.CF$_UFI_CODE_A1149 from IFSAPP.A1135_INVPART_CHAR_ALL_CFV ch where ch.CONTRACT = 'AD' and ch.PART_NO = '" + _PART_NO + "' AND ch.CHARACTERISTIC_CODE = 'W001'";
                string sUFI_CODE = ff.Oracle_GetValue(sConString, sSQL_UFI, ref sErr);
                sLine = "UFI_CODE = " + sUFI_CODE;
                sw.WriteLine(sLine);

                //
                if (checkBoxDETONATOR_APLIC.Checked) // doplněno 22.9.2014 Senkerik
                {
                    sLine = "DETONATOR_APPLIC = NO";
                    sw.WriteLine(sLine);
                }
                // 2.4.2019
                if (sCatalog_desc.ToUpper().Contains("DUAL"))
                {
                    sLine = "WELD_APPLIC = NO"; //vypínač zatavení
                    sw.WriteLine(sLine);
                    sLine = "FLAG_APPLIC = NO"; //vypínač praporku
                    sw.WriteLine(sLine);
                }
                if (!sCatalog_desc.ToUpper().Contains("DUAL") && checkBoxFLAG_APPLIC.Checked)
                {
                    sLine = "FLAG_APPLIC = NO"; //vypínač praporku
                    sw.WriteLine(sLine);
                }

                checkBox_AlterCK.Checked = false;

                //string sTime = String.Format("{0}-{1}-{2} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                string sSQLp = "INSERT INTO nela(ZO_BatchNo,VZ_BatchNo,Cesta,DatumExportu,Uzivatel,Nela) VALUES(" +
                            "'" + _ZO_BatchNo + "'," +
                            "'" + _VZ_BatchNo + "'," +
                            "'" + sFile.Trim() + "'," +
                            "'" + sTime + "','" + _UserName.Trim() + "'," +
                            "'" + _NELA.Trim() + "'" +
                            ")";
                //(String.Format("yyyy-MM-dd HH:mm:ss",DateTime.Now) + ("\'" + ")")))))))))))));
                cmdO.Dispose();
                cnO.Close();
                if (MSSQL_Execute(sSQLp) == false)
                {
                    MessageBox.Show("fce: MSSQL_Execute neproběhla v pořádku.");
                }
                Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
                cn.Close();
                sw.Dispose();
                sw.Close();
                MessageBox.Show("Soubor je vytvořen.");

                KopirujNalepky(true);
            }
            catch (Exception ex)
            {
                if (sw != null)
                {
                    sw.Close();
                }
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }
            this.Cursor = Cursors.Default;
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fldOpen = new FolderBrowserDialog();
            fldOpen.SelectedPath = textBoxFolderToSave.Text;

            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            if (fldOpen.ShowDialog() == DialogResult.OK)
            {
                // Assign the cursor in the Stream to the Form's Cursor property.
                textBoxFolderToSave.Text = fldOpen.SelectedPath;

                Globals.FOLDER_BTW_EXPORT = fldOpen.SelectedPath + @"\export";

                Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                config.AppSettings.Settings["OutputFolder"].Value = textBoxFolderToSave.Text;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");

                //Properties.Settings.Default["OutputFolder"] = fldOpen.SelectedPath;
                //Properties.Settings.Default.Save();

            }
            fldOpen.Dispose();

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
#if DEBUG
                string applicationName = Environment.GetCommandLineArgs()[0];
#else
           string applicationName = Environment.GetCommandLineArgs()[0];
#endif
                string exePath = System.IO.Path.Combine(Environment.CurrentDirectory, applicationName);
                System.Configuration.Configuration configNew = ConfigurationManager.OpenExeConfiguration(exePath);


                //ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
                //fileMap.ExeConfigFilename = @Application.ProductName.ToString() + ".exe.config";
                //Configuration config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

                if (ConfigurationManager.AppSettings != null)
                {
                    if (configNew.AppSettings.Settings.Count > 0)
                    {
                        configNew.AppSettings.Settings["OutputFolder"].Value = textBoxFolderToSave.Text;

                        configNew.Save(ConfigurationSaveMode.Full);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }

        }
        private bool MSSQL_Execute(string pSQL)
        {
            string cnsSQL01 = "Persist Security Info=True;User ID=epo;Password=123456;Initial Catalog=epo;Data Source=SQL01";

            SqlConnection cnSQL01;
            cnSQL01 = new SqlConnection(cnsSQL01);
            cnSQL01.Open();
            SqlTransaction trx;
            trx = cnSQL01.BeginTransaction("VlozDoSQL");
            SqlCommand cmd;
            cmd = new SqlCommand(pSQL, cnSQL01);
            try
            {
                //trx = cnSQL01.BeginTransaction("VlozDoSQL");
                cmd.Transaction = trx;
                cmd.ExecuteNonQuery();
                trx.Commit();
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("VlozDoSQL-" + ex.Message.ToString());
                trx.Rollback();
                return false;
            }
            finally
            {
                cmd.Dispose();
            }
        }
        private string MSSQL_Najdi(string pSQL)
        {
            string cnsSQL01 = "Persist Security Info=True;User ID=epo;Password=123456;Initial Catalog=epo;Data Source=SQL01";
            SqlConnection cn;
            try
            {
                cn = new SqlConnection(cnsSQL01);
                SqlCommand cmd;
                cn.Open();
                cmd = new SqlCommand(pSQL, cn);
                string sPom = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();
                cn.Close();
                return sPom;
            }
            catch (Exception ex)
            {
                MessageBox.Show("MSSQL_Najdi-" + ex.Message.ToString());
                return "";
            }
            finally
            {
            }
        }
        private string MSSQL_GetOneValue(string pConnectString, string pSQL)
        {
            string cnsSQL01 = pConnectString;
            string sPom = string.Empty;

            SqlConnection cn;
            try
            {
                cn = new SqlConnection(cnsSQL01);
                SqlCommand cmd;
                cn.Open();
                cmd = new SqlCommand(pSQL, cn);
                sPom = Convert.ToString(cmd.ExecuteScalar());
                cmd.Dispose();
                cn.Close();
                return sPom;
            }
            catch (Exception ex)
            {
                MessageBox.Show("MSSQL_Najdi-" + ex.Message.ToString());
                return string.Empty;
            }
            finally
            {
            }
        }
        private void Napln_dataGridViewNalepky()
        {
            this.Cursor = Cursors.WaitCursor;

            string sSQLp = string.Empty;

           System.Data.OracleClient.OracleConnection cn = default(System.Data.OracleClient.OracleConnection);
            //string sCnString = "Data Source=ADPROD;UID=userbarco;PWD=userbarco";
            string sCnString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;";
            System.Data.OracleClient.OracleDataAdapter adapter = new System.Data.OracleClient.OracleDataAdapter();

            _dsNalepky = new DataSet();
            // bere z c_labels_ad, ale nedá PRO
            //string sSQLp = "SELECT distinct file_name FROM IFSAPP.c_labels_ad a" + " WHERE a.order_no = '" + _Customer_order + "'" + 
            //    " AND  a.line_no = '" + _LINE_NO + "'" + 
            //    " AND  a.release_no = '" + _RELEASE_NO + "'" + 
            //    " AND  a.line_item_no = '" + _LINE_ITEM_NO + "'" + 
            //    " ORDER BY a.file_name";


            // brát vše z DMS (z ftp)
            if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
            {
                sSQLp = Globals.SQL_Nalepky_From_DMS_IFS9 + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'";

            }
            else
            {
                sSQLp = Globals.SQL_Nalepky_From_PO + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'";

            }

            cn = new System.Data.OracleClient.OracleConnection(sCnString);
            cn.Open();

            System.Data.OracleClient.OracleCommand cmd = new System.Data.OracleClient.OracleCommand(sSQLp, cn);
            //
            adapter.SelectCommand = cmd;
            adapter.Fill(_dsNalepky, "Nalepky");
            DataTable dt = _dsNalepky.Tables[0];
            //dgZakazky.DisplayMember = dt.Columns("doc_no").ToString
            dataGridViewNalepky.BindingContext = new BindingContext();
            dataGridViewNalepky.DataSource = dt;
            //
            dataGridViewNalepky.Columns["FILE_NAME"].HeaderText = "FILE_NAME";
            dataGridViewNalepky.Columns["FILE_NAME"].Width = 300;

            //CheckedListBox1.DisplayMember = "doc_no"
            cmd.Dispose();
            cn.Close();
            if (dataGridViewNalepky.CurrentRow != null)
            {
                //textBoxNalepka.Text = "\\\\Bap01\\IFS_NALEPKY\\" + this.dataGridViewNalepky[this.dataGridViewNalepky.CurrentRow.Index, 0].Value.ToString();
                if (Globals.TNSNAME == TNS_Names.AD9PROD)
                { textBoxNalepka.Text = Globals.FOLDER_NALEPKY_IFS9PROD + this.dataGridViewNalepky[this.dataGridViewNalepky.CurrentRow.Index, 0].Value.ToString(); }
                else
                { textBoxNalepka.Text = Globals.FOLDER_NALEPKY_IFS9TEST + this.dataGridViewNalepky[this.dataGridViewNalepky.CurrentRow.Index, 0].Value.ToString(); }
                    
            }
            this.Cursor = Cursors.Default;
        }

        private void buttonNaplnNalepky_Click(object sender, EventArgs e)
        {
            Napln_dataGridViewNalepky();
        }

        private void dataGridViewVZ_CurrentCellChanged(object sender, EventArgs e)
        {
        }

        private void dataGridViewVZ_SelectionChanged(object sender, EventArgs e)
        {

            string sKey = null;

            if (dataGridViewVZ.CurrentRow == null)
            {
                return;
            }
            _Current_Row = dataGridViewVZ.CurrentRow.Index;

            dataGridViewNalepky.DataSource = null;
            sKey = dataGridViewVZ[0, this.dataGridViewVZ.CurrentRow.Index].Value.ToString();

            System.Data.DataRow[] foundRows = null;
            foundRows = _DS.Tables[0].Select("VZ_BatchNo ='" + sKey + "'");

            _VZ_Order_No = Convert.ToString(dataGridViewVZ.CurrentRow.Cells[2].Value);
            _VZ_Release_No = Convert.ToString(dataGridViewVZ.CurrentRow.Cells[3].Value);
            _VZ_sequence_No = Convert.ToString(dataGridViewVZ.CurrentRow.Cells[4].Value);

            //mDs.Tables(0).
            _Customer_order = foundRows[0][7].ToString();
            ZO.ORDER_NO = _Customer_order;

            _LINE_NO = foundRows[0][8].ToString();
            ZO.LINE_NO = _LINE_NO;

            _RELEASE_NO = foundRows[0][9].ToString();
            ZO.RELEASE_NO = _RELEASE_NO;

            _LINE_ITEM_NO = foundRows[0][10].ToString();
            ZO.LINE_ITEM_NO = _LINE_ITEM_NO;

            ZO.VZ_BatchNo = foundRows[0][0].ToString();
            ZO.ZO_BatchNo = foundRows[0][1].ToString();
            ZO.PART_NO = foundRows[0][5].ToString();
            ZO.Customer = foundRows[0][13].ToString();

            Napln_dataGridViewNalepky();

            foundRows = null;
        }

        private void dataGridViewNalepky_SelectionChanged(object sender, EventArgs e)
        {
            _FILE_NAME = dataGridViewNalepky[0, this.dataGridViewNalepky.CurrentRow.Index].Value.ToString();
            if (Globals.TNSNAME == TNS_Names.AD9PROD)
            {
                textBoxNalepka.Text = Globals.FOLDER_NALEPKY_IFS9PROD + _FILE_NAME;
            }
            else
            { textBoxNalepka.Text = Globals.FOLDER_NALEPKY_IFS9TEST+ _FILE_NAME; }
            

        }

        private void dataGridViewVZ_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            _COLUMN_INDEX = e.ColumnIndex;
            _COLUMN_NAME = dataGridViewVZ.Columns[_COLUMN_INDEX].Name;
            _COLUMN_NAME = dataGridViewVZ.Columns[_COLUMN_INDEX].HeaderText;
            labelHledaSePodle.Text = "Hledá se podle :" + _COLUMN_NAME;
        }

        private void buttonFind_Click(object sender, EventArgs e)
        {
            HledejVz();
        }
        private void HledejVz()
        {
#pragma warning disable CS0168 // The variable 'bs' is declared but never used
            BindingContext bs;
#pragma warning restore CS0168 // The variable 'bs' is declared but never used
            // sKey = dgZakazky.Item(Me.dgZakazky.CurrentRowIndex(), 0)
            // myDataTable.Select("VZ_BatchNo ='" + sKey + "'")
#pragma warning disable CS0168 // The variable 'foundRows' is declared but never used
            DataRow[] foundRows;
#pragma warning restore CS0168 // The variable 'foundRows' is declared but never used
            if (_DS == null)
            {
                return;
            }
            DataTable dt = _DS.Tables[0];
            int iC = dt.Rows.Count;
            int i;
            int iAkt = dataGridViewVZ.CurrentRow.Index;
            string sPom = string.Empty;

            for (i = 0; (i <= (iC - 1)); i++)
            {
                sPom = dataGridViewVZ[_COLUMN_INDEX, i].Value.ToString().Trim();
                //sPom = sPom.Substring(0, textBoxFindVZ.Text.Length);
                if ((sPom.Contains(textBoxFindVZ.Text.Trim())))
                {
                    dataGridViewVZ.BindingContext[_DS.Tables[0]].Position = i;
                    dataGridViewVZ.Rows[iAkt].Selected = false;
                    dataGridViewVZ.Rows[i].Selected = true;
                    return;
                }
            }
            // foundRows = 
            // mDs.Tables(0).Select("VZ_BatchNo LIKE '" + tbHledejVz.Text + "%'")
            // dgZakazky.CurrentRowIndex() 
            // dgZakazky.BindingContext(mDs.Tables(0)).Position = mDs.Tables(0).Select("VZ_BatchNo LIKE '" + tbHledejVz.Text + "%'"))
            // BindingSource1.Position = BindingSource1.Find(DS.Tables(0).Columns(FieldNameToSearch).ToString, KeysPressed)

            //
            for (int j = 0; j < (dataGridViewVZ.Rows.Count); j++)
            {
                if (dataGridViewVZ.Rows[j].Cells[_COLUMN_INDEX].Value.ToString().StartsWith(textBoxFindVZ.Text, true, CultureInfo.CurrentUICulture))
                {
                    dataGridViewVZ.Rows[j].Selected = true;
                    //_Index_HV_OperaceOsoby = i;
                    dataGridViewVZ.FirstDisplayedScrollingRowIndex = j;
                }
            }

            //
        }

        private void buttonKopirujNalepky_Click(object sender, EventArgs e)
        {
            KopirujNalepky(true);
        }
        private void KopirujNalepky(bool pNotify)
        {
            // IFS/Správa dokumentů to ukládá přímo zde   S:\IT_IFS_DOC_MAN\AD9PROD\NALEPKY\ (\\austin.cz\cz\AD\SHARE\IT_IFS_DOC_MAN\AD9PROD\NALEPKY\)
            // a tak se to bere přímo odsud ne z BAP01.
            string sPom = string.Empty;
            //
            if (textBoxFolderToSave.Text.ToString().Trim() == string.Empty || !Directory.Exists(textBoxFolderToSave.Text.ToString().Trim()))
            {
                MessageBox.Show("Není zadána cesta (nebo neexistuje), pro ukládání nálepek !", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBoxFolderToSave.Focus();
                textBoxFolderToSave.SelectAll();
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            try
            {
                DataTable dt = (DataTable)dataGridViewNalepky.DataSource;
                string sNalepka = string.Empty;
                string sPathLocal = Application.StartupPath + @"\export\";
                string sCesta = textBoxFolderToSave.Text.ToString() + @"\Export\";
                if (Directory.Exists(sCesta))
                {
                    //System.IO.Directory.Delete(sCesta, true);
                }
                else
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(textBoxFolderToSave.Text.ToString());
                    DirectorySecurity dirSec = dirInfo.GetAccessControl();
                    dirSec.AddAccessRule(new FileSystemAccessRule("AUSTIN\\" + _UserName, FileSystemRights.Modify, AccessCo‌ntrolType.Allow));
                    dirInfo.SetAccessControl(dirSec);

                    System.IO.Directory.CreateDirectory(sCesta, dirSec);
                }

                if (dt == null)
                {
                    MessageBox.Show("dataGridViewNalepky je prázdný !");
                    return;
                }
                foreach (DataRow row in dt.Rows)
                {
                    sNalepka = row[0].ToString().Trim();
                    if (sNalepka.Contains("-KTN-") || sNalepka.Contains("-STN-") || sNalepka.Contains("-PRO-")) // || sNalepka.Contains("-SBS-")
                    {
                        //continue;
                    }
                    else
                    {
                        continue;
                    }
                    //sPom = sNalepka;

                    //if (File.Exists(Globals.FOLDER_NALEPKY + sNalepka))
                    //{
                    //    System.IO.File.Copy(Globals.FOLDER_NALEPKY + sNalepka, sCesta + sNalepka, true);
                    //}

                    // ftp
                    string sErr = string.Empty;


                    //  \\austin.cz\cz\AD\SHARE\IT_IFS_DOC_MAN\AD9UPG3\NALEPKY
                    string sFileNalepka = string.Empty;

                    if (Globals.TNSNAME == TNS_Names.AD9PROD)
                    {
                        sFileNalepka = Globals.FOLDER_NALEPKY_IFS9PROD + sNalepka;
                    }
                    else
                    {
                        sFileNalepka = Globals.FOLDER_NALEPKY_IFS9TEST + sNalepka;
                    }

                    //if (Globals.TNSNAME == TNS_Names.AD9PROD)
                    //{
                        if (File.Exists(sFileNalepka))
                        {
                            if (!GrantAccess(sCesta + sNalepka, sCesta))
                            { return; }
                            if (File.Exists(sCesta + sNalepka))
                            {
                                File.SetAttributes(sCesta + sNalepka, FileAttributes.Normal);
                                File.Delete(sCesta + sNalepka);
                            }
                            File.Copy(sFileNalepka, sCesta + sNalepka, true);
                            try
                            {
                                if (GrantAccess(sCesta + sNalepka, sCesta))
                                {
                                    File.SetAttributes(sCesta + sNalepka, FileAttributes.Normal);
                                    File.SetCreationTime(sCesta + sNalepka, DateTime.Now);
                                    File.SetLastWriteTime(sCesta + sNalepka, DateTime.Now);
                                    File.SetLastAccessTime(sCesta + sNalepka, DateTime.Now);
                                }
                            }
                            catch (Exception ex1)
                            {
                                MessageBox.Show(ex1.Message.ToString());
                            }
                        }
                        else
                        {
                            frmMessage f = new frmMessage();
                            f.Message = "Zkontrolujte zda soubor existuje v IFS (není na úložišti)!!" + Environment.NewLine +
                                sFileNalepka;
                            f.ShowDialog();
                            f.Dispose();
                            //MessageBox.Show((Globals.FOLDER_NALEPKY_IFS9 + sNalepka + Environment.NewLine + Environment.NewLine +
                            //    "Zkontrolujte zda soubor existuje v IFS (není na úložišti)!!"), "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                //    }
                //    else
                //    {
                //        if (!GrantAccess(sCesta + sNalepka, sCesta))
                //        { return; }
                //        if (File.Exists(sCesta + sNalepka))
                //        {
                //            File.SetAttributes(sCesta + sNalepka, FileAttributes.Normal);
                //            File.Delete(sCesta + sNalepka);
                //        }
                //        if (!Ftp_CopyFile(sNalepka, sCesta + sNalepka, ref sErr))
                //        {
                //            //MessageBox.Show(sErr, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                //            // kontrola existence složky
                //            if (!Directory.Exists(sPathLocal))
                //            { 
                //                try
                //                {
                //                    Directory.CreateDirectory(sPathLocal);
                //                }
                //                catch (Exception exx)
                //                {
                //                    MessageBox.Show(exx.Message.ToString());
                //                    continue;
                //                }
                //            }
                //            // kopíruju z ftp na cečko a potom do požadované složky zadané v programu
                //            if (!Ftp_CopyFile(sNalepka, sPathLocal + sNalepka, ref sErr))
                //            {
                //                MessageBox.Show(sNalepka + Environment.NewLine + Environment.NewLine +
                //                    "Zkontrolujte zda soubor existuje v IFS (není na ftp)!!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //            }
                //            else
                //            {
                //                // copy do zadané složky
                //                File.Copy(sPathLocal + sNalepka, sCesta + sNalepka, true);
                //                try
                //                {
                //                    if (GrantAccess(sCesta + sNalepka, sCesta))
                //                    {
                //                        File.SetAttributes(sCesta + sNalepka, FileAttributes.Normal);
                //                        File.SetCreationTime(sCesta + sNalepka, DateTime.Now);
                //                        File.SetLastWriteTime(sCesta + sNalepka, DateTime.Now);
                //                        File.SetLastAccessTime(sCesta + sNalepka, DateTime.Now);
                //                    }
                //                }
                //                catch (Exception ex2)
                //                {
                //                    MessageBox.Show(ex2.Message.ToString());
                //                }
                //            }

                //        }
                //        else
                //        {
                //        }
                //    }
                //    //FtpClient ftp1 = new FtpClient();
                //    //if (ftp1.ftp_download("BAP01", "IFSFTP", "IfSftP2Ad", "Nalepky",sNalepka,sCesta+sNalepka) == false)
                //    //    {
                //    //        MessageBox.Show("Chyba","Nepodařilo se zkopírovat nálepku : " + Environment.NewLine + sNalepka
                //    //   ,MessageBoxButtons.OK,MessageBoxIcon.Error);

                //    //    ftp1.Close();
                //    //    return;
                //    //}
                //    //ftp1.Close();
                }
                if (pNotify)
                { 
                    MessageBox.Show("Nálepky zkopírovány do :" + sCesta);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Ex:" + Environment.NewLine + ex.Message);
            }
            this.Cursor = Cursors.Default;
        }

        private void KopirujTXTaNalepky_NA_IMT_FTP(bool pNotify)
        {
            string sPom = string.Empty;
            int iPocetZkopirovanychNalepek = 0;

            if (!NaplnListBtw())
            {
                return;
            }

            //
            if (textBoxFolderToSave.Text.ToString().Trim() == string.Empty || !Directory.Exists(textBoxFolderToSave.Text.ToString().Trim()))
            {
                MessageBox.Show("Není zadána cesta (nebo neexistuje), pro ukládání vygenerovaného souboru !", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBoxFolderToSave.Focus();
                textBoxFolderToSave.SelectAll();
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // copy TXT
                FtpClient ftp = new FtpClient();
                if (ftp.ftp_file_exist(Globals.IMT_FTP_SERVER_IMT, "imt", "imt123456", "MEMCARD1/RecipesADR", Globals.IMT_FILE_RECEPTURA) == true)
                {
                    MessageBox.Show("Soubor už je ve schválených");
                    this.Cursor = Cursors.Default;
                    return;
                }
                if (ftp.ftp_copy(Globals.IMT_FTP_SERVER, "austin", "austin123456", "zakazky", Globals.IMT_FOLDERFILE_RECEPTURA) == false)
                {
                    MessageBox.Show("Kopírování TXT (receptury) na FTP se nezdařilo");
                    this.Cursor = Cursors.Default;
                    return;
                }
                else
                {
                    MessageBox.Show("Soubor: " + Environment.NewLine + Globals.IMT_FILE_RECEPTURA + Environment.NewLine +
                        "byl zkopírován na IMT - FTP.", "FTP - COPY");

                }
                ftp.Close();
                //DataTable dt = (DataTable)dataGridViewNalepky.DataSource;
                string sNalepka = string.Empty;
                //string sPathLocal = Application.StartupPath + @"\export\";
                string sCesta = Globals.FOLDER_NALEPKY; // textBoxFolderToSave.Text.ToString() + @"\Export\";
                if (Directory.Exists(sCesta))
                {
                }
                else
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(textBoxFolderToSave.Text.ToString());
                    DirectorySecurity dirSec = dirInfo.GetAccessControl();
                    dirSec.AddAccessRule(new FileSystemAccessRule("AUSTIN\\" + _UserName, FileSystemRights.Modify, AccessCo‌ntrolType.Allow));
                    dirInfo.SetAccessControl(dirSec);
                    MessageBox.Show("Problém s právy na složku, nebo neexistuje :" + sCesta, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                    //System.IO.Directory.CreateDirectory(sCesta, dirSec);
                }

                FtpClient ftp1 = new FtpClient();
                foreach (string item in _Labels_BTW_names)
                {
                    if (!System.IO.File.Exists(sCesta + item))
                    {
                        MessageBox.Show("Soubor:" + sCesta + item + Environment.NewLine +" neexistuje !!");
                    }
                    else
                    {  
                        if (ftp1.ftp_copy(Globals.IMT_FTP_SERVER, "austin", "austin123456", "nalepky", sCesta + item) == false)
                        {
                            MessageBox.Show("Kopírování souboru na FTP se nezdařilo", sCesta + sNalepka);
                        }
                        else
                        {
                            iPocetZkopirovanychNalepek++;
                        }
                    }
                }
                ftp1.Close();
                if (pNotify)
                {
                    MessageBox.Show("Nálepky zkopírovány na FTP - IMT" + Environment.NewLine + "Počet : "+ _Labels_BTW_names.Count.ToString() + "/" + iPocetZkopirovanychNalepek.ToString());
                }

                // zapis do logu
                string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                string sSQLp = "Update nela SET DatumNaImtFtp = " +
                            "'" + sTime + "'" +
                            " WHERE ID =" + Globals.ID_TABLE_NELA;

                //(String.Format("yyyy-MM-dd HH:mm:ss",DateTime.Now) + ("\'" + ")")))))))))))));
                if (MSSQL_Execute(sSQLp) == false)
                {
                    MessageBox.Show("fce: MSSQL_Execute neproběhla v pořádku.");
                }
                Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);            //

            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Ex:" + Environment.NewLine + ex.Message);
            }
            this.Cursor = Cursors.Default;
        }

        private void buttonSort_Click(object sender, EventArgs e)
        {
            List<DataGridViewColumn> dgvColumns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn dgvCol in dataGridViewVZ.Columns)
                dgvColumns.Add(dgvCol);

            //Pass the Column list to the sort Form
            using (SortColumnsForm form = new SortColumnsForm(this, dgvColumns))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    GridRowComparer rowComparer = new GridRowComparer(form.SortOrderList);
                    for (int i = 0; i < form.SortOrderList.Count - 1; i++)
                    {
                        string sColumn = form.SortOrderList[0].Key.DataPropertyName;
                        this.dataGridViewVZ.Sort(this.dataGridViewVZ.Columns[sColumn], ListSortDirection.Ascending);
                    }

                    //this.dataGridViewVZ.Sort(this.dataGridView1.Columns["Name"], ListSortDirection.Ascending);

                    //dataGridViewVZ.Sort(rowComparer);
                }
            }

        }

        internal void PerformSort(List<KeyValuePair<DataGridViewColumn, bool>> sortOrderList)
        {
            GridRowComparer rowComparer = new GridRowComparer(sortOrderList);
            dataGridViewVZ.Sort(rowComparer);
        }

        private void ToolStripMenuItem_otevřiTXT_Click(object sender, EventArgs e)
        {

            if (dataGridViewSoubory.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte řádek");
                return;
            }


            DataGridViewRow currentRow = dataGridViewSoubory.SelectedRows[0];

            string sFileName = Convert.ToString(currentRow.Cells["Cesta"].Value);
            if (!File.Exists(sFileName))
            {
                MessageBox.Show("Soubor neexistuje" + Environment.NewLine + sFileName, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //
            frmFileView f = new frmFileView();
            f.FileName = sFileName;
            f.ShowDialog();
            f.Dispose();

            //
            //Process.Start("notepad.exe", sFileName);
        }
        private void ToolStripMenuItemOpravitTXT_Click(object sender, EventArgs e)
        {
            if (dataGridViewSoubory.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte řádek");
                return;
            }


            DataGridViewRow currentRow = dataGridViewSoubory.SelectedRows[0];

            string sFileName = Convert.ToString(currentRow.Cells["Cesta"].Value);
            if (!File.Exists(sFileName))
            {
                MessageBox.Show("Soubor neexistuje" + Environment.NewLine + sFileName, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //
            Process.Start("notepad.exe", sFileName);
        }
        private void ProgramyAD_Log_Zapis()
        {
            this.Cursor = Cursors.WaitCursor;
            string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            DateTime dt = DateTime.Now;
            sTime = dt.ToString();
            string sSQLp = "INSERT INTO ProgramyAD_Log(DATUM,COMPUTER,PROGRAM,VERSION,UZIVATEL,IPADDRESS) VALUES(" +
                        "'" + sTime + "'," +
                        "'" + _ComputerName + "'," +
                        "'" + Application.ProductName + "'," +
                        "'" + _Version + "'," +
                        "'" + _UserName + "'," + "'" + this._IPAddress + "')";

            Funkce ff = new Funkce();
            ff.ConString_MSSQL_Austin_Trc = Globals.ConString_MSSQL_Austin_Trc;
            string sVloz = ff.MSSQL_Execute(sSQLp);

            if (sVloz != "")
            {
                MessageBox.Show("Nepodařilo se vložit záznam do AUSTIN_TRC.ProgramyAD_LOG," + Environment.NewLine +
                    " Volejte správce 4185 nebo 733 611 525 !", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
            }
            ff = null;
            this.Cursor = Cursors.Default;
        }
        private bool Ftp_CopyFile(string pFileName, string pDestinationFolderFile, ref string pErr)
        {
            string sFtpFolder = string.Empty;
            try
            {
                using (WebClient ftpClient = new WebClient())
                {
                    if (Globals.TNSNAME.Contains("UPG"))
                    {
                        sFtpFolder = "ftp://IFSDB/NALEPKY/";
                        ftpClient.Credentials = new System.Net.NetworkCredential("ftpuser", "ftpUser");
                    }
                    else
                    {
                        sFtpFolder = "ftp://bap01/NALEPKY/";
                        ftpClient.Credentials = new System.Net.NetworkCredential("IFSFTP", "IfSftP2Ad");
                    }

                    ftpClient.DownloadFile(sFtpFolder + pFileName, pDestinationFolderFile);
                    ftpClient.Dispose();
                }
                return true;
            }
            catch (Exception ex)
            {
                pErr = "Exception message : " + ex.Message + (char)13 +
                    "Cíl včetně cesty : " + pDestinationFolderFile + (char)13 +
                    "Ze složky na ftp : " + sFtpFolder;
                //Console.WriteLine(ex.InnerException.ToString());
                return false;
            }
        }

        private void buttonMojeNalepky_Click(object sender, EventArgs e)
        {
            if (dataGridViewVZ.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte řádek");
                return;
            }
            //_NELA = comboBoxNelaCislo.SelectedItem.ToString().Trim();

            DataGridViewRow currentRow = dataGridViewVZ.SelectedRows[0];
            _VZ_BatchNo = Convert.ToString(currentRow.Cells[0].Value);
            _ZO_BatchNo = Convert.ToString(currentRow.Cells[1].Value);
            _Customer_order = Convert.ToString(currentRow.Cells[7].Value);
            _LINE_NO = Convert.ToString(currentRow.Cells[8].Value);
            _RELEASE_NO = Convert.ToString(currentRow.Cells[9].Value);
            _LINE_ITEM_NO = Convert.ToString(currentRow.Cells[10].Value);
            _PART_NO = Convert.ToString(currentRow.Cells[5].Value);

            comboBoxNelaCislo.SelectedIndex = 0;

            VytvorSouborProDA();
            //

        }
        private void MakeDataTable_Nalepky()
        // Create a new DataTable named NewProducts. 
        {
            //c01_DOC_CLASS
            //c02_TITLE
            //c03_DOC_NO
            //c04_DOC_REV
            //c05_FILE_NAME
            //c06_CUSTOMER_PO_NO
            //c07_ORDER_NO
            //c08_LINE_NO
            //c09_RELEASE_NO
            //c10_LINE_ITEM_NO
            //c11_PACKAGE_NUM1
            //c12_SEQ_NUM
            //c13_PACKAGE_NUM2
            //c14_CATALOG_NO
            //c15_CATALOG_DESC
            //c16_VENDOR_COMPANY
            //c17_VENDOR_NAME
            //c18_VENDOR_ADDRESS_ID
            //c19_VENDOR_STREET
            //c20_VENDOR_CITY
            //c21_VENDOR_ZIP_CODE
            //c22_VENDOR_COUNTRY
            //c23_RECEIVER_COMPANY_NAME2
            //c24_RECEIVER_ADDRESS1
            //c25_RECEIVER_ADDRESS2
            //c26_RECEIVER_ZIP_CODE
            //c27_RECEIVER_CITY
            //c28_RECEIVER_COUNTRY
            //c29_RECEIVER_NAME
            //c30_RECEIVER_PHONE
            //c31_RECEIVER_FAX
            //c32_RECEIVER_MOBILE
            //c33_RECEIVER_EMAIL
            //c34_UNIQUE_PACKAGE_ID
            //c35_UNIQUE_SUBPACKAGE_ID
            //c36_PACKAGE_QTY
            //c37_SMALL_PACKAGE_QTY
            //c38_PACKAGE_GROSS_WEIGHT
            //c39_ROUND_NET_WEIGHT
            //c40_PACKAGE_NET_WEIGHT
            //c41_EXPLOSIVE_NET_WEIGHT
            //c42_UN_NO
            //c43_ADR_RID_CLASS
            //c44_DETONATOR_CLASS
            //c45_DELAY_LEVEL
            //c46_CAVITY_MATERIAL
            //c47_IGNITION_CUE_MJ
            //c48_IGNITION_CUE_J
            //c49_SAFETY_STREAM
            //c50_SAFETY_CUE_MJ
            //c51_SAFETY_CUE_J
            //c52_SUPPLY_WIRE_MATERIAL
            //c53_WIRE_ISULATOR_MATERIAL
            //c54_WIRE_LENGTH
            //c55_WIRE_LENGTH_TRACE
            //c56_CONSUMPTION_DATE
            //c57_NOMINAL_TIME_EA
            //c58_NOMINAL_TIME_NA
            //c59_ST_LENGTH
            //c60_DELAY_STEP
            //c61_BATCH_NO
            //c62_TESTING_STATION_NUMBER
            //c63_TESTING_PROTOCOL_NUMBER
            //c64_CUSTOMER_AUTHORIZE_DECREE
            //c65_ANOTHER_NOTICE
            //c66_NOTE
            //c67_OBJVERSION
            //c68_OBJID
            //c69_DATE
            //c70_EXPIRATION
            //c71_TOTAL_RESIST
            //c72_BRIDGE_RESIST
            //c73_DATE_CODE
            //c74_LOT_NUMBER
            //c75_STATION_NUM
            //c76_WORKERS
            //c77_WORKER2
            //c78_DUPLICATE
            //c79_PIECE_NUM
            //c80_DELAY_FRACTION
            //c81_GENERAL
            //c82_GENERAL_1
            //c83_BOX_PCS
            //c84_MANUAL
            //c85_PRODUCER
            //c86_DMX_PACKAGE_QTY
            //c87_DMX_DELAY_LEVEL
            //c88_DMX_DELAY_STEP
            //c89_DMX_NOMINAL_TIME_EA
            //c90_DMX_NOMINAL_TIME_NA
            //c91_DMX_WIRE_LENGTH
            //c92_DMX_ST_LENGTH
            //c93_DMX_DATE
            //c94_DMX_PACKAGE_NUM2
            //c95_DMX_SEQ_NUM
            //c96_DMX_PRODUCER
            //c97_AUTO_LABEL
            //c98_TOTAL_PACKAGE_QTY

            //DataTable newProducts = new DataTable("NewProducts");
            _dtNalepky = new DataTable();
            _keysNalepky = new DataColumn[1];

            // Add three column objects to the table. 
            //keys[0] = Station;
            // pomocný pro rozlišení typu nálepky
            DataColumn BTW_TYPE = new DataColumn(); BTW_TYPE.DataType = System.Type.GetType("System.String"); BTW_TYPE.ColumnName = btw.BTW_TYPE; _dtNalepky.Columns.Add(BTW_TYPE);

            DataColumn c01_DOC_CLASS = new DataColumn(); c01_DOC_CLASS.DataType = System.Type.GetType("System.String"); c01_DOC_CLASS.ColumnName = "c01_DOC_CLASS"; _dtNalepky.Columns.Add(c01_DOC_CLASS);
            DataColumn c02_TITLE = new DataColumn(); c02_TITLE.DataType = System.Type.GetType("System.String"); c02_TITLE.ColumnName = "c02_TITLE"; _dtNalepky.Columns.Add(c02_TITLE);
            DataColumn c03_DOC_NO = new DataColumn(); c03_DOC_NO.DataType = System.Type.GetType("System.String"); c03_DOC_NO.ColumnName = "c03_DOC_NO"; _dtNalepky.Columns.Add(c03_DOC_NO);
            DataColumn c04_DOC_REV = new DataColumn(); c04_DOC_REV.DataType = System.Type.GetType("System.String"); c04_DOC_REV.ColumnName = "c04_DOC_REV"; _dtNalepky.Columns.Add(c04_DOC_REV);
            DataColumn c05_FILE_NAME = new DataColumn(); c05_FILE_NAME.DataType = System.Type.GetType("System.String"); c05_FILE_NAME.ColumnName = "c05_FILE_NAME"; _dtNalepky.Columns.Add(c05_FILE_NAME);
            DataColumn c06_CUSTOMER_PO_NO = new DataColumn(); c06_CUSTOMER_PO_NO.DataType = System.Type.GetType("System.String"); c06_CUSTOMER_PO_NO.ColumnName = "c06_CUSTOMER_PO_NO"; _dtNalepky.Columns.Add(c06_CUSTOMER_PO_NO);
            DataColumn c07_ORDER_NO = new DataColumn(); c07_ORDER_NO.DataType = System.Type.GetType("System.String"); c07_ORDER_NO.ColumnName = "c07_ORDER_NO"; _dtNalepky.Columns.Add(c07_ORDER_NO);
            DataColumn c08_LINE_NO = new DataColumn(); c08_LINE_NO.DataType = System.Type.GetType("System.String"); c08_LINE_NO.ColumnName = "c08_LINE_NO"; _dtNalepky.Columns.Add(c08_LINE_NO);
            DataColumn c09_RELEASE_NO = new DataColumn(); c09_RELEASE_NO.DataType = System.Type.GetType("System.String"); c09_RELEASE_NO.ColumnName = "c09_RELEASE_NO"; _dtNalepky.Columns.Add(c09_RELEASE_NO);
            DataColumn c10_LINE_ITEM_NO = new DataColumn(); c10_LINE_ITEM_NO.DataType = System.Type.GetType("System.String"); c10_LINE_ITEM_NO.ColumnName = "c10_LINE_ITEM_NO"; _dtNalepky.Columns.Add(c10_LINE_ITEM_NO);
            DataColumn c11_PACKAGE_NUM1 = new DataColumn(); c11_PACKAGE_NUM1.DataType = System.Type.GetType("System.String"); c11_PACKAGE_NUM1.ColumnName = "c11_PACKAGE_NUM1"; _dtNalepky.Columns.Add(c11_PACKAGE_NUM1);
            DataColumn c12_SEQ_NUM = new DataColumn(); c12_SEQ_NUM.DataType = System.Type.GetType("System.String"); c12_SEQ_NUM.ColumnName = "c12_SEQ_NUM"; _dtNalepky.Columns.Add(c12_SEQ_NUM);
            DataColumn c13_PACKAGE_NUM2 = new DataColumn(); c13_PACKAGE_NUM2.DataType = System.Type.GetType("System.String"); c13_PACKAGE_NUM2.ColumnName = "c13_PACKAGE_NUM2"; _dtNalepky.Columns.Add(c13_PACKAGE_NUM2);
            DataColumn c14_CATALOG_NO = new DataColumn(); c14_CATALOG_NO.DataType = System.Type.GetType("System.String"); c14_CATALOG_NO.ColumnName = "c14_CATALOG_NO"; _dtNalepky.Columns.Add(c14_CATALOG_NO);
            DataColumn c15_CATALOG_DESC = new DataColumn(); c15_CATALOG_DESC.DataType = System.Type.GetType("System.String"); c15_CATALOG_DESC.ColumnName = "c15_CATALOG_DESC"; _dtNalepky.Columns.Add(c15_CATALOG_DESC);
            DataColumn c16_VENDOR_COMPANY = new DataColumn(); c16_VENDOR_COMPANY.DataType = System.Type.GetType("System.String"); c16_VENDOR_COMPANY.ColumnName = "c16_VENDOR_COMPANY"; _dtNalepky.Columns.Add(c16_VENDOR_COMPANY);
            DataColumn c17_VENDOR_NAME = new DataColumn(); c17_VENDOR_NAME.DataType = System.Type.GetType("System.String"); c17_VENDOR_NAME.ColumnName = "c17_VENDOR_NAME"; _dtNalepky.Columns.Add(c17_VENDOR_NAME);
            DataColumn c18_VENDOR_ADDRESS_ID = new DataColumn(); c18_VENDOR_ADDRESS_ID.DataType = System.Type.GetType("System.String"); c18_VENDOR_ADDRESS_ID.ColumnName = "c18_VENDOR_ADDRESS_ID"; _dtNalepky.Columns.Add(c18_VENDOR_ADDRESS_ID);
            DataColumn c19_VENDOR_STREET = new DataColumn(); c19_VENDOR_STREET.DataType = System.Type.GetType("System.String"); c19_VENDOR_STREET.ColumnName = "c19_VENDOR_STREET"; _dtNalepky.Columns.Add(c19_VENDOR_STREET);
            DataColumn c20_VENDOR_CITY = new DataColumn(); c20_VENDOR_CITY.DataType = System.Type.GetType("System.String"); c20_VENDOR_CITY.ColumnName = "c20_VENDOR_CITY"; _dtNalepky.Columns.Add(c20_VENDOR_CITY);
            DataColumn c21_VENDOR_ZIP_CODE = new DataColumn(); c21_VENDOR_ZIP_CODE.DataType = System.Type.GetType("System.String"); c21_VENDOR_ZIP_CODE.ColumnName = "c21_VENDOR_ZIP_CODE"; _dtNalepky.Columns.Add(c21_VENDOR_ZIP_CODE);
            DataColumn c22_VENDOR_COUNTRY = new DataColumn(); c22_VENDOR_COUNTRY.DataType = System.Type.GetType("System.String"); c22_VENDOR_COUNTRY.ColumnName = "c22_VENDOR_COUNTRY"; _dtNalepky.Columns.Add(c22_VENDOR_COUNTRY);
            DataColumn c23_RECEIVER_COMPANY_NAME2 = new DataColumn(); c23_RECEIVER_COMPANY_NAME2.DataType = System.Type.GetType("System.String"); c23_RECEIVER_COMPANY_NAME2.ColumnName = "c23_RECEIVER_COMPANY_NAME2"; _dtNalepky.Columns.Add(c23_RECEIVER_COMPANY_NAME2);
            DataColumn c24_RECEIVER_ADDRESS1 = new DataColumn(); c24_RECEIVER_ADDRESS1.DataType = System.Type.GetType("System.String"); c24_RECEIVER_ADDRESS1.ColumnName = "c24_RECEIVER_ADDRESS1"; _dtNalepky.Columns.Add(c24_RECEIVER_ADDRESS1);
            DataColumn c25_RECEIVER_ADDRESS2 = new DataColumn(); c25_RECEIVER_ADDRESS2.DataType = System.Type.GetType("System.String"); c25_RECEIVER_ADDRESS2.ColumnName = "c25_RECEIVER_ADDRESS2"; _dtNalepky.Columns.Add(c25_RECEIVER_ADDRESS2);
            DataColumn c26_RECEIVER_ZIP_CODE = new DataColumn(); c26_RECEIVER_ZIP_CODE.DataType = System.Type.GetType("System.String"); c26_RECEIVER_ZIP_CODE.ColumnName = "c26_RECEIVER_ZIP_CODE"; _dtNalepky.Columns.Add(c26_RECEIVER_ZIP_CODE);
            DataColumn c27_RECEIVER_CITY = new DataColumn(); c27_RECEIVER_CITY.DataType = System.Type.GetType("System.String"); c27_RECEIVER_CITY.ColumnName = "c27_RECEIVER_CITY"; _dtNalepky.Columns.Add(c27_RECEIVER_CITY);
            DataColumn c28_RECEIVER_COUNTRY = new DataColumn(); c28_RECEIVER_COUNTRY.DataType = System.Type.GetType("System.String"); c28_RECEIVER_COUNTRY.ColumnName = "c28_RECEIVER_COUNTRY"; _dtNalepky.Columns.Add(c28_RECEIVER_COUNTRY);
            DataColumn c29_RECEIVER_NAME = new DataColumn(); c29_RECEIVER_NAME.DataType = System.Type.GetType("System.String"); c29_RECEIVER_NAME.ColumnName = "c29_RECEIVER_NAME"; _dtNalepky.Columns.Add(c29_RECEIVER_NAME);
            DataColumn c30_RECEIVER_PHONE = new DataColumn(); c30_RECEIVER_PHONE.DataType = System.Type.GetType("System.String"); c30_RECEIVER_PHONE.ColumnName = "c30_RECEIVER_PHONE"; _dtNalepky.Columns.Add(c30_RECEIVER_PHONE);
            DataColumn c31_RECEIVER_FAX = new DataColumn(); c31_RECEIVER_FAX.DataType = System.Type.GetType("System.String"); c31_RECEIVER_FAX.ColumnName = "c31_RECEIVER_FAX"; _dtNalepky.Columns.Add(c31_RECEIVER_FAX);
            DataColumn c32_RECEIVER_MOBILE = new DataColumn(); c32_RECEIVER_MOBILE.DataType = System.Type.GetType("System.String"); c32_RECEIVER_MOBILE.ColumnName = "c32_RECEIVER_MOBILE"; _dtNalepky.Columns.Add(c32_RECEIVER_MOBILE);
            DataColumn c33_RECEIVER_EMAIL = new DataColumn(); c33_RECEIVER_EMAIL.DataType = System.Type.GetType("System.String"); c33_RECEIVER_EMAIL.ColumnName = "c33_RECEIVER_EMAIL"; _dtNalepky.Columns.Add(c33_RECEIVER_EMAIL);
            DataColumn c34_UNIQUE_PACKAGE_ID = new DataColumn(); c34_UNIQUE_PACKAGE_ID.DataType = System.Type.GetType("System.String"); c34_UNIQUE_PACKAGE_ID.ColumnName = "c34_UNIQUE_PACKAGE_ID"; _dtNalepky.Columns.Add(c34_UNIQUE_PACKAGE_ID);
            DataColumn c35_UNIQUE_SUBPACKAGE_ID = new DataColumn(); c35_UNIQUE_SUBPACKAGE_ID.DataType = System.Type.GetType("System.String"); c35_UNIQUE_SUBPACKAGE_ID.ColumnName = "c35_UNIQUE_SUBPACKAGE_ID"; _dtNalepky.Columns.Add(c35_UNIQUE_SUBPACKAGE_ID);
            DataColumn c36_PACKAGE_QTY = new DataColumn(); c36_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c36_PACKAGE_QTY.ColumnName = "c36_PACKAGE_QTY"; _dtNalepky.Columns.Add(c36_PACKAGE_QTY);
            DataColumn c37_SMALL_PACKAGE_QTY = new DataColumn(); c37_SMALL_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c37_SMALL_PACKAGE_QTY.ColumnName = "c37_SMALL_PACKAGE_QTY"; _dtNalepky.Columns.Add(c37_SMALL_PACKAGE_QTY);
            DataColumn c38_PACKAGE_GROSS_WEIGHT = new DataColumn(); c38_PACKAGE_GROSS_WEIGHT.DataType = System.Type.GetType("System.String"); c38_PACKAGE_GROSS_WEIGHT.ColumnName = "c38_PACKAGE_GROSS_WEIGHT"; _dtNalepky.Columns.Add(c38_PACKAGE_GROSS_WEIGHT);
            DataColumn c39_ROUND_NET_WEIGHT = new DataColumn(); c39_ROUND_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c39_ROUND_NET_WEIGHT.ColumnName = "c39_ROUND_NET_WEIGHT"; _dtNalepky.Columns.Add(c39_ROUND_NET_WEIGHT);
            DataColumn c40_PACKAGE_NET_WEIGHT = new DataColumn(); c40_PACKAGE_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c40_PACKAGE_NET_WEIGHT.ColumnName = "c40_PACKAGE_NET_WEIGHT"; _dtNalepky.Columns.Add(c40_PACKAGE_NET_WEIGHT);
            DataColumn c41_EXPLOSIVE_NET_WEIGHT = new DataColumn(); c41_EXPLOSIVE_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c41_EXPLOSIVE_NET_WEIGHT.ColumnName = "c41_EXPLOSIVE_NET_WEIGHT"; _dtNalepky.Columns.Add(c41_EXPLOSIVE_NET_WEIGHT);
            DataColumn c42_UN_NO = new DataColumn(); c42_UN_NO.DataType = System.Type.GetType("System.String"); c42_UN_NO.ColumnName = "c42_UN_NO"; _dtNalepky.Columns.Add(c42_UN_NO);
            DataColumn c43_ADR_RID_CLASS = new DataColumn(); c43_ADR_RID_CLASS.DataType = System.Type.GetType("System.String"); c43_ADR_RID_CLASS.ColumnName = "c43_ADR_RID_CLASS"; _dtNalepky.Columns.Add(c43_ADR_RID_CLASS);
            DataColumn c44_DETONATOR_CLASS = new DataColumn(); c44_DETONATOR_CLASS.DataType = System.Type.GetType("System.String"); c44_DETONATOR_CLASS.ColumnName = "c44_DETONATOR_CLASS"; _dtNalepky.Columns.Add(c44_DETONATOR_CLASS);
            DataColumn c45_DELAY_LEVEL = new DataColumn(); c45_DELAY_LEVEL.DataType = System.Type.GetType("System.String"); c45_DELAY_LEVEL.ColumnName = "c45_DELAY_LEVEL"; _dtNalepky.Columns.Add(c45_DELAY_LEVEL);
            DataColumn c46_CAVITY_MATERIAL = new DataColumn(); c46_CAVITY_MATERIAL.DataType = System.Type.GetType("System.String"); c46_CAVITY_MATERIAL.ColumnName = "c46_CAVITY_MATERIAL"; _dtNalepky.Columns.Add(c46_CAVITY_MATERIAL);
            DataColumn c47_IGNITION_CUE_MJ = new DataColumn(); c47_IGNITION_CUE_MJ.DataType = System.Type.GetType("System.String"); c47_IGNITION_CUE_MJ.ColumnName = "c47_IGNITION_CUE_MJ"; _dtNalepky.Columns.Add(c47_IGNITION_CUE_MJ);
            DataColumn c48_IGNITION_CUE_J = new DataColumn(); c48_IGNITION_CUE_J.DataType = System.Type.GetType("System.String"); c48_IGNITION_CUE_J.ColumnName = "c48_IGNITION_CUE_J"; _dtNalepky.Columns.Add(c48_IGNITION_CUE_J);
            DataColumn c49_SAFETY_STREAM = new DataColumn(); c49_SAFETY_STREAM.DataType = System.Type.GetType("System.String"); c49_SAFETY_STREAM.ColumnName = "c49_SAFETY_STREAM"; _dtNalepky.Columns.Add(c49_SAFETY_STREAM);
            DataColumn c50_SAFETY_CUE_MJ = new DataColumn(); c50_SAFETY_CUE_MJ.DataType = System.Type.GetType("System.String"); c50_SAFETY_CUE_MJ.ColumnName = "c50_SAFETY_CUE_MJ"; _dtNalepky.Columns.Add(c50_SAFETY_CUE_MJ);
            DataColumn c51_SAFETY_CUE_J = new DataColumn(); c51_SAFETY_CUE_J.DataType = System.Type.GetType("System.String"); c51_SAFETY_CUE_J.ColumnName = "c51_SAFETY_CUE_J"; _dtNalepky.Columns.Add(c51_SAFETY_CUE_J);
            DataColumn c52_SUPPLY_WIRE_MATERIAL = new DataColumn(); c52_SUPPLY_WIRE_MATERIAL.DataType = System.Type.GetType("System.String"); c52_SUPPLY_WIRE_MATERIAL.ColumnName = "c52_SUPPLY_WIRE_MATERIAL"; _dtNalepky.Columns.Add(c52_SUPPLY_WIRE_MATERIAL);
            DataColumn c53_WIRE_ISULATOR_MATERIAL = new DataColumn(); c53_WIRE_ISULATOR_MATERIAL.DataType = System.Type.GetType("System.String"); c53_WIRE_ISULATOR_MATERIAL.ColumnName = "c53_WIRE_ISULATOR_MATERIAL"; _dtNalepky.Columns.Add(c53_WIRE_ISULATOR_MATERIAL);
            DataColumn c54_WIRE_LENGTH = new DataColumn(); c54_WIRE_LENGTH.DataType = System.Type.GetType("System.String"); c54_WIRE_LENGTH.ColumnName = "c54_WIRE_LENGTH"; _dtNalepky.Columns.Add(c54_WIRE_LENGTH);
            DataColumn c55_WIRE_LENGTH_TRACE = new DataColumn(); c55_WIRE_LENGTH_TRACE.DataType = System.Type.GetType("System.String"); c55_WIRE_LENGTH_TRACE.ColumnName = "c55_WIRE_LENGTH_TRACE"; _dtNalepky.Columns.Add(c55_WIRE_LENGTH_TRACE);
            DataColumn c56_CONSUMPTION_DATE = new DataColumn(); c56_CONSUMPTION_DATE.DataType = System.Type.GetType("System.String"); c56_CONSUMPTION_DATE.ColumnName = "c56_CONSUMPTION_DATE"; _dtNalepky.Columns.Add(c56_CONSUMPTION_DATE);
            DataColumn c57_NOMINAL_TIME_EA = new DataColumn(); c57_NOMINAL_TIME_EA.DataType = System.Type.GetType("System.String"); c57_NOMINAL_TIME_EA.ColumnName = "c57_NOMINAL_TIME_EA"; _dtNalepky.Columns.Add(c57_NOMINAL_TIME_EA);
            DataColumn c58_NOMINAL_TIME_NA = new DataColumn(); c58_NOMINAL_TIME_NA.DataType = System.Type.GetType("System.String"); c58_NOMINAL_TIME_NA.ColumnName = "c58_NOMINAL_TIME_NA"; _dtNalepky.Columns.Add(c58_NOMINAL_TIME_NA);
            DataColumn c59_ST_LENGTH = new DataColumn(); c59_ST_LENGTH.DataType = System.Type.GetType("System.String"); c59_ST_LENGTH.ColumnName = "c59_ST_LENGTH"; _dtNalepky.Columns.Add(c59_ST_LENGTH);
            DataColumn c60_DELAY_STEP = new DataColumn(); c60_DELAY_STEP.DataType = System.Type.GetType("System.String"); c60_DELAY_STEP.ColumnName = "c60_DELAY_STEP"; _dtNalepky.Columns.Add(c60_DELAY_STEP);
            DataColumn c61_BATCH_NO = new DataColumn(); c61_BATCH_NO.DataType = System.Type.GetType("System.String"); c61_BATCH_NO.ColumnName = "c61_BATCH_NO"; _dtNalepky.Columns.Add(c61_BATCH_NO);
            DataColumn c62_TESTING_STATION_NUMBER = new DataColumn(); c62_TESTING_STATION_NUMBER.DataType = System.Type.GetType("System.String"); c62_TESTING_STATION_NUMBER.ColumnName = "c62_TESTING_STATION_NUMBER"; _dtNalepky.Columns.Add(c62_TESTING_STATION_NUMBER);
            DataColumn c63_TESTING_PROTOCOL_NUMBER = new DataColumn(); c63_TESTING_PROTOCOL_NUMBER.DataType = System.Type.GetType("System.String"); c63_TESTING_PROTOCOL_NUMBER.ColumnName = "c63_TESTING_PROTOCOL_NUMBER"; _dtNalepky.Columns.Add(c63_TESTING_PROTOCOL_NUMBER);
            DataColumn c64_CUSTOMER_AUTHORIZE_DECREE = new DataColumn(); c64_CUSTOMER_AUTHORIZE_DECREE.DataType = System.Type.GetType("System.String"); c64_CUSTOMER_AUTHORIZE_DECREE.ColumnName = "c64_CUSTOMER_AUTHORIZE_DECREE"; _dtNalepky.Columns.Add(c64_CUSTOMER_AUTHORIZE_DECREE);
            DataColumn c65_ANOTHER_NOTICE = new DataColumn(); c65_ANOTHER_NOTICE.DataType = System.Type.GetType("System.String"); c65_ANOTHER_NOTICE.ColumnName = "c65_ANOTHER_NOTICE"; _dtNalepky.Columns.Add(c65_ANOTHER_NOTICE);
            DataColumn c66_NOTE = new DataColumn(); c66_NOTE.DataType = System.Type.GetType("System.String"); c66_NOTE.ColumnName = "c66_NOTE"; _dtNalepky.Columns.Add(c66_NOTE);
            DataColumn c67_OBJVERSION = new DataColumn(); c67_OBJVERSION.DataType = System.Type.GetType("System.String"); c67_OBJVERSION.ColumnName = "c67_OBJVERSION"; _dtNalepky.Columns.Add(c67_OBJVERSION);
            DataColumn c68_OBJID = new DataColumn(); c68_OBJID.DataType = System.Type.GetType("System.String"); c68_OBJID.ColumnName = "c68_OBJID"; _dtNalepky.Columns.Add(c68_OBJID);
            DataColumn c69_DATE = new DataColumn(); c69_DATE.DataType = System.Type.GetType("System.String"); c69_DATE.ColumnName = "c69_DATE"; _dtNalepky.Columns.Add(c69_DATE);
            DataColumn c70_EXPIRATION = new DataColumn(); c70_EXPIRATION.DataType = System.Type.GetType("System.String"); c70_EXPIRATION.ColumnName = "c70_EXPIRATION"; _dtNalepky.Columns.Add(c70_EXPIRATION);
            DataColumn c71_TOTAL_RESIST = new DataColumn(); c71_TOTAL_RESIST.DataType = System.Type.GetType("System.String"); c71_TOTAL_RESIST.ColumnName = "c71_TOTAL_RESIST"; _dtNalepky.Columns.Add(c71_TOTAL_RESIST);
            DataColumn c72_BRIDGE_RESIST = new DataColumn(); c72_BRIDGE_RESIST.DataType = System.Type.GetType("System.String"); c72_BRIDGE_RESIST.ColumnName = "c72_BRIDGE_RESIST"; _dtNalepky.Columns.Add(c72_BRIDGE_RESIST);
            DataColumn c73_DATE_CODE = new DataColumn(); c73_DATE_CODE.DataType = System.Type.GetType("System.String"); c73_DATE_CODE.ColumnName = "c73_DATE_CODE"; _dtNalepky.Columns.Add(c73_DATE_CODE);
            DataColumn c74_LOT_NUMBER = new DataColumn(); c74_LOT_NUMBER.DataType = System.Type.GetType("System.String"); c74_LOT_NUMBER.ColumnName = "c74_LOT_NUMBER"; _dtNalepky.Columns.Add(c74_LOT_NUMBER);
            DataColumn c75_STATION_NUM = new DataColumn(); c75_STATION_NUM.DataType = System.Type.GetType("System.String"); c75_STATION_NUM.ColumnName = "c75_STATION_NUM"; _dtNalepky.Columns.Add(c75_STATION_NUM);
            DataColumn c76_WORKERS = new DataColumn(); c76_WORKERS.DataType = System.Type.GetType("System.String"); c76_WORKERS.ColumnName = "c76_WORKERS"; _dtNalepky.Columns.Add(c76_WORKERS);
            DataColumn c77_WORKER2 = new DataColumn(); c77_WORKER2.DataType = System.Type.GetType("System.String"); c77_WORKER2.ColumnName = "c77_WORKER2"; _dtNalepky.Columns.Add(c77_WORKER2);
            DataColumn c78_DUPLICATE = new DataColumn(); c78_DUPLICATE.DataType = System.Type.GetType("System.String"); c78_DUPLICATE.ColumnName = "c78_DUPLICATE"; _dtNalepky.Columns.Add(c78_DUPLICATE);
            DataColumn c79_PIECE_NUM = new DataColumn(); c79_PIECE_NUM.DataType = System.Type.GetType("System.String"); c79_PIECE_NUM.ColumnName = "c79_PIECE_NUM"; _dtNalepky.Columns.Add(c79_PIECE_NUM);
            DataColumn c80_DELAY_FRACTION = new DataColumn(); c80_DELAY_FRACTION.DataType = System.Type.GetType("System.String"); c80_DELAY_FRACTION.ColumnName = "c80_DELAY_FRACTION"; _dtNalepky.Columns.Add(c80_DELAY_FRACTION);
            DataColumn c81_GENERAL = new DataColumn(); c81_GENERAL.DataType = System.Type.GetType("System.String"); c81_GENERAL.ColumnName = "c81_GENERAL"; _dtNalepky.Columns.Add(c81_GENERAL);
            DataColumn c82_GENERAL_1 = new DataColumn(); c82_GENERAL_1.DataType = System.Type.GetType("System.String"); c82_GENERAL_1.ColumnName = "c82_GENERAL_1"; _dtNalepky.Columns.Add(c82_GENERAL_1);
            DataColumn c83_BOX_PCS = new DataColumn(); c83_BOX_PCS.DataType = System.Type.GetType("System.String"); c83_BOX_PCS.ColumnName = "c83_BOX_PCS"; _dtNalepky.Columns.Add(c83_BOX_PCS);
            DataColumn c84_MANUAL = new DataColumn(); c84_MANUAL.DataType = System.Type.GetType("System.String"); c84_MANUAL.ColumnName = "c84_MANUAL"; _dtNalepky.Columns.Add(c84_MANUAL);
            DataColumn c85_PRODUCER = new DataColumn(); c85_PRODUCER.DataType = System.Type.GetType("System.String"); c85_PRODUCER.ColumnName = "c85_PRODUCER"; _dtNalepky.Columns.Add(c85_PRODUCER);
            DataColumn c86_DMX_PACKAGE_QTY = new DataColumn(); c86_DMX_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c86_DMX_PACKAGE_QTY.ColumnName = "c86_DMX_PACKAGE_QTY"; _dtNalepky.Columns.Add(c86_DMX_PACKAGE_QTY);
            DataColumn c87_DMX_DELAY_LEVEL = new DataColumn(); c87_DMX_DELAY_LEVEL.DataType = System.Type.GetType("System.String"); c87_DMX_DELAY_LEVEL.ColumnName = "c87_DMX_DELAY_LEVEL"; _dtNalepky.Columns.Add(c87_DMX_DELAY_LEVEL);
            DataColumn c88_DMX_DELAY_STEP = new DataColumn(); c88_DMX_DELAY_STEP.DataType = System.Type.GetType("System.String"); c88_DMX_DELAY_STEP.ColumnName = "c88_DMX_DELAY_STEP"; _dtNalepky.Columns.Add(c88_DMX_DELAY_STEP);
            DataColumn c89_DMX_NOMINAL_TIME_EA = new DataColumn(); c89_DMX_NOMINAL_TIME_EA.DataType = System.Type.GetType("System.String"); c89_DMX_NOMINAL_TIME_EA.ColumnName = "c89_DMX_NOMINAL_TIME_EA"; _dtNalepky.Columns.Add(c89_DMX_NOMINAL_TIME_EA);
            DataColumn c90_DMX_NOMINAL_TIME_NA = new DataColumn(); c90_DMX_NOMINAL_TIME_NA.DataType = System.Type.GetType("System.String"); c90_DMX_NOMINAL_TIME_NA.ColumnName = "c90_DMX_NOMINAL_TIME_NA"; _dtNalepky.Columns.Add(c90_DMX_NOMINAL_TIME_NA);
            DataColumn c91_DMX_WIRE_LENGTH = new DataColumn(); c91_DMX_WIRE_LENGTH.DataType = System.Type.GetType("System.String"); c91_DMX_WIRE_LENGTH.ColumnName = "c91_DMX_WIRE_LENGTH"; _dtNalepky.Columns.Add(c91_DMX_WIRE_LENGTH);
            DataColumn c92_DMX_ST_LENGTH = new DataColumn(); c92_DMX_ST_LENGTH.DataType = System.Type.GetType("System.String"); c92_DMX_ST_LENGTH.ColumnName = "c92_DMX_ST_LENGTH"; _dtNalepky.Columns.Add(c92_DMX_ST_LENGTH);
            DataColumn c93_DMX_DATE = new DataColumn(); c93_DMX_DATE.DataType = System.Type.GetType("System.String"); c93_DMX_DATE.ColumnName = "c93_DMX_DATE"; _dtNalepky.Columns.Add(c93_DMX_DATE);
            DataColumn c94_DMX_PACKAGE_NUM2 = new DataColumn(); c94_DMX_PACKAGE_NUM2.DataType = System.Type.GetType("System.String"); c94_DMX_PACKAGE_NUM2.ColumnName = "c94_DMX_PACKAGE_NUM2"; _dtNalepky.Columns.Add(c94_DMX_PACKAGE_NUM2);
            DataColumn c95_DMX_SEQ_NUM = new DataColumn(); c95_DMX_SEQ_NUM.DataType = System.Type.GetType("System.String"); c95_DMX_SEQ_NUM.ColumnName = "c95_DMX_SEQ_NUM"; _dtNalepky.Columns.Add(c95_DMX_SEQ_NUM);
            DataColumn c96_DMX_PRODUCER = new DataColumn(); c96_DMX_PRODUCER.DataType = System.Type.GetType("System.String"); c96_DMX_PRODUCER.ColumnName = "c96_DMX_PRODUCER"; _dtNalepky.Columns.Add(c96_DMX_PRODUCER);
            DataColumn c97_AUTO_LABEL = new DataColumn(); c97_AUTO_LABEL.DataType = System.Type.GetType("System.String"); c97_AUTO_LABEL.ColumnName = "c97_AUTO_LABEL"; _dtNalepky.Columns.Add(c97_AUTO_LABEL);
            DataColumn c98_TOTAL_PACKAGE_QTY = new DataColumn(); c98_TOTAL_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c98_TOTAL_PACKAGE_QTY.ColumnName = "c98_TOTAL_PACKAGE_QTY"; _dtNalepky.Columns.Add(c98_TOTAL_PACKAGE_QTY);

        }

        private void VytvorSouborProDA()
        {
            this.Cursor = Cursors.WaitCursor;
            //
            string sFILENAME_KTN = string.Empty;
            string sFILENAME_KTN2 = string.Empty;
            string sFILENAME_STN = string.Empty;
            string sFILENAME_PRO = string.Empty;
            string sFILENAME_PRO_PRIPONA = string.Empty;
            string sTitle = string.Empty;
            //
            MakeDataTable_Nalepky();

            string sConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;"; ;
            //OracleConnection cn;
            StreamWriter sw;

            string sFilePreprog = (_NELA == NelaFile_MachineName.E3PR ? Globals.FILE_E3PR_TXT : Globals.FILE_DA001_TXT);

            string sLine;
            string sWhereOrderKomplet = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'";
            string sWhereOrderLineKomplet = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'" +
                            " AND  line_item_no = " + _LINE_ITEM_NO;

            string sWhereOrderLineKompletForSalesPart = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'" +
                            " AND  line_item_no <= 0";

            string sBatchNo = _ZO_BatchNo;
            string sPACKAGE_QTY_MULTI = string.Empty;
            string sNOMINAL_TIME_NA_MULTI = string.Empty;
            string sDelay_step_multi = string.Empty;
            string sAUTO_LABEL = string.Empty;
            string sAUTO_LABEL2 = string.Empty;

            string sFile = (textBoxFolderToSave.Text.Trim() + ("\\" + (_VZ_BatchNo + ".TXT")));
            if (_NELA == NelaFile_MachineName.E3PR)
            {
                sFile = (textBoxFolderToSave.Text.Trim() + ("\\" + (_ZO_BatchNo + ".TXT")));
            }
            Globals.IMT_FILE_RECEPTURA = _VZ_BatchNo + ".TXT";
            Globals.IMT_FOLDERFILE_RECEPTURA = sFile;
            sw = new StreamWriter(sFile, false, System.Text.Encoding.Default);
            try
            {
                int i = 0;
                string sErr = string.Empty;
                string sSQL_1 = string.Empty;
                Funkce ff = new Funkce();
                // 
                System.Data.OracleClient.OracleConnection cn = new System.Data.OracleClient.OracleConnection(sConString);
                cn.Open();
                // adresy vendor a receiver
                // CUSTOMER_INFO_ADDRESS.ADDRESS2
                sSQL_1 = Globals.SQL_GetIDZakaznika.Replace("xxORDER_NO", _Customer_order);
                string sCustomer_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = Globals.SQL_GetIDAddress.Replace("xxORDER_NO", _Customer_order);
                string sIDAddress = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                //public static string SQL_GetAddress = "SELECT xxColumn  FROM IFSAPP.CUSTOMER_INFO_ADDRESS a where a.customer_id = 'xxCustomer_id' AND a.address_id = 'xxAddress_id'";
                sSQL_1 = Globals.SQL_GetAddress.Replace("xxColumn", "a.ADDRESS1");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);

                string sReceiverAddress1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.ADDRESS1", "a.ADDRESS2");
                string sReceiverAddress2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.ADDRESS2", "a.CITY");
                string sReceiverCity = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.CITY", "a.COUNTRY");
                string sReceiverCountry = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.COUNTRY", "a.ZIP_CODE");
                string sReceiverZIP_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCompanyNameZakaznika.Replace("xxCustomer_id", sCustomer_no);
                string sReceiverCompanyName = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD.Replace("xxColumn", "a.Name");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactName = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD.Replace("xxColumn", "a.value");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactPhone = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_EMAIL.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactEmail = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_FAX.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactFax = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_MOBILE.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactMobile = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);


                // VENDOR
                //
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = Globals.SQL_Get_xml_production_site_number;
                }
                else
                {
                    sSQL_1 = Globals.SQL_Get_ProducerCode_Prefix;
                }

                string sPRODUCER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_ADDRESS_ID
                sSQL_1 = Globals.SQL_GetVendor_Address_ID;
                string sVENDOR_ADDRESS_ID = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_CITY
                sSQL_1 = Globals.SQL_GetVendor_City;
                string sVENDOR_CITY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COMPANY = 'AD'
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_COMPANY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COUNTRY
                sSQL_1 = Globals.SQL_Get_Language_Code_FromOrder.Replace("xxORDER_NO", _Customer_order);
                string sLanguageCode = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = Globals.SQL_GetVendor_Country.Replace("xxLANGUAGE_CODE", sLanguageCode);
                string sVENDOR_COUNTRY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_NAME
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_NAME = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_STREET

                //
                //btw.VENDOR_ZIP_CODE



                // jedná se o úsporné balení ?
                string sSQL_UsporneBaleni = "SELECT a.order_no FROM  IFSAPP.C_PACKAGE_HEAD a" +
                            " WHERE a.unique_pkg = 'FALSE'" + // lišky z IFS to daly opačně !!! když je false tak se jedná o úsporné balení
                            " AND  a.order_no = '" + _Customer_order + "'" +
                            " AND  a.line_no = '" + _LINE_NO + "'" +
                            " AND  a.rel_no = '" + _RELEASE_NO + "'";

                bool bUsporneBaleni = ff.Oracle_Exists(sConString, sSQL_UsporneBaleni, ref sErr);
                if (bUsporneBaleni)
                {
                    string sSQL_Multi = "SELECT * FROM (SELECT distinct order_no,package_num2,part_no," +
                        "replace(wm_concat(amount) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') PACKAGE_QTY_MULTI," +
                        "replace(wm_concat(nominal_time) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') NOMINAL_TIME_NA_MULTI," +
                        "replace(wm_concat(delay_step_multi) OVER (PARTITION BY package_num2 ORDER BY order_no,package_num2) ,',','+') Delay_step_multi" +
                        " FROM" +
                        "( " +
                        "   SELECT a.order_no,b.package_num2,b.ch_line_item_no,b.amount,b.part_no," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'N016') as nominal_time," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'L004') as delay_step_multi" +
                        "   FROM  IFSAPP.C_PACKAGE_HEAD a, IFSAPP.C_PACKAGE_DETAIL b " +
                        "   WHERE a.order_no = b.order_no" +
                        "   AND a.line_no = b.line_no" +
                        "   AND a.rel_no = b.rel_no" +
                        "   AND a.line_item_no = b.line_item_no" +
                        "   AND a.unique_pkg = 'FALSE'" +
                        "   AND  a.order_no = '" + _Customer_order + "'" +
                        "   AND  a.line_no = '" + _LINE_NO + "'" +
                        "   AND  a.rel_no = '" + _RELEASE_NO + "'" +
                        "   AND  a.line_item_no = '" + _LINE_ITEM_NO + "'" +
                        "   ORDER BY order_no,ch_line_item_no" +
                        "))" +
                        " WHERE part_no = '" + _PART_NO + "'" +
                        " ORDER BY order_no,package_num2";
                    // při úsporném balení vrátí více řádků (kartonů) pro výrobu jedné položky 
                    sPACKAGE_QTY_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "PACKAGE_QTY_MULTI", ref sErr);
                    sNOMINAL_TIME_NA_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "NOMINAL_TIME_NA_MULTI", ref sErr);
                    sDelay_step_multi = ff.Oracle_GetValue(sConString, sSQL_Multi, "delay_step_multi", ref sErr);


                }

                //
                sLine = "[TABLE]";
                sw.WriteLine(sLine);
                sLine = "MACHINE_NAME = " + _NELA;
                sw.WriteLine(sLine);
                sLine = "STATION_NUM = " + _NELA;
                sw.WriteLine(sLine);
                //PRODUCER
                sLine = "PRODUCER = " + sPRODUCER;
                sw.WriteLine(sLine);
                //ORDER_NO
                sLine = "ORDER_NO = " + _Customer_order;
                sw.WriteLine(sLine);
                //CUSTOMER_PO_NO 
                sSQL_1 = "SELECT CUSTOMER_PO_NO FROM IFSAPP.CUSTOMER_ORDER WHERE order_no = '" + _Customer_order + "'";
                string sCUSTOMER_PO_NO = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CUSTOMER_PO_NO = " + sCUSTOMER_PO_NO;
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                {
                    sw.WriteLine(sLine);
                }


                //LINE_NO
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                {
                    sLine = "LINE_NO = " + _LINE_NO;
                    sw.WriteLine(sLine);
                }

                //RELEASE_NO
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                {
                    sLine = "RELEASE_NO = " + _RELEASE_NO;
                    sw.WriteLine(sLine);
                }


                //LINE_ITEM_NO
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                {
                    sLine = "LINE_ITEM_NO = " + _LINE_ITEM_NO;
                    sw.WriteLine(sLine);
                }

                //BATCH_NO
                sLine = ("BATCH_NO = " + sBatchNo);
                sw.WriteLine(sLine);

                if (_Preprog)
                {
                    sLine = ("PREPROG = " + "PREPROG_" + sBatchNo + ".TXT");
                }
                else sLine = "PREPROG = ";
                sw.WriteLine(sLine);


                // KARTONY
                //string sSQL_K = ("SELECT distinct file_name FROM IFSAPP.c_labels_ad a" + (" WHERE a.order_no = \'"
                //            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                //            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                //            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                //            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%KTN%'" + " ORDER BY a.file_name"))))))))))))));

                // výběr kartonů  ze zakázek z IFS 
                string sSQL_K = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", _Customer_order);
                sSQL_K = sSQL_K.Replace("xxLINE_NO", _LINE_NO);
                sSQL_K = sSQL_K.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_K = sSQL_K.Replace("xxTypNalepky", "KTN");
                System.Data.OracleClient.OracleCommand cmdK = new System.Data.OracleClient.OracleCommand(sSQL_K, cn);
                System.Data.OracleClient.OracleDataReader rdK = cmdK.ExecuteReader();
                if (!rdK.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-KTN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                        sw.WriteLine(sLine);
                    }
                }
                while (rdK.Read())
                {
                    i++;
                    if ((i == 2))
                    {
                        if (rdK[0].GetType() != typeof(DBNull))
                        {
                            sLine = ("FILENAME_KTN2 = " + rdK.GetString(0).Trim());
                        }
                        else
                        {
                            sLine = "FILENAME_KTN2 = ";
                        }
                    }
                    else if ((rdK[0].GetType() != typeof(DBNull)))
                    {
                        sLine = ("FILENAME_KTN = " + rdK.GetString(0).Trim());
                        sFILENAME_KTN = rdK.GetString(0).Trim();
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                    }
                    sw.WriteLine(sLine);
                }
                if ((i == 1))
                {
                    sLine = "FILENAME_KTN2 =";
                    sw.WriteLine(sLine);
                }
                rdK.Close();
                cmdK.Dispose();


                //  pro Skladacku
                string sHodnotaPole = string.Empty;
                //string sSQL_S = ("SELECT batch_no,doc_class,MAX(file_name) AS FileName FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                //            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                //            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                //            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                //            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%STN%'" + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));

                string sSQL_S = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", _Customer_order);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", _LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "STN");

                System.Data.OracleClient.OracleCommand cmdS = new System.Data.OracleClient.OracleCommand(sSQL_S, cn);
                System.Data.OracleClient.OracleDataReader rdS = cmdS.ExecuteReader();
                if (!rdS.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-STN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_STN = ";
                        sw.WriteLine(sLine);
                    }
                }
                // FILENAME_STN
                sLine = "FILENAME_STN = ";
                int iPocetSTN = 0;
                while (rdS.Read())
                {
                    iPocetSTN++;
                    if ((iPocetSTN == 2))
                    {
                        if (rdS[0].GetType() != typeof(DBNull))
                        {
                            sLine = ("FILENAME_STN2 = " + rdS.GetString(0).Trim());
                            sHodnotaPole = rdS["FileName"].ToString();
                            sHodnotaPole = sHodnotaPole.Replace("-A1.BTW", "");
                            sAUTO_LABEL2 = MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'");
                            sAUTO_LABEL2 = "AUTO_LABEL = " + sAUTO_LABEL2;

                        }
                        else
                        {
                            sLine = "FILENAME_STN2 = ";
                        }
                    }
                    else if ((rdS["FileName"].GetType() == typeof(DBNull)))
                    {
                        sLine = "FILENAME_STN = ";
                    }
                    else
                    {
                        sLine = ("FILENAME_STN = " + rdS["FileName"].ToString());
                        sFILENAME_STN = rdS["FileName"].ToString();
                        //AUTO_LABEL
                        sHodnotaPole = rdS["FileName"].ToString();
                        sHodnotaPole = sHodnotaPole.Replace("-A1.BTW", "");
                        sAUTO_LABEL = MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'");
                        sAUTO_LABEL = "AUTO_LABEL = " + sAUTO_LABEL;

                    }
                    sw.WriteLine(sLine);
                    if (rdS["Title"].GetType() != typeof(DBNull))
                    {
                        sTitle = rdS["Title"].ToString();
                    }
                }
                if (iPocetSTN == 1)
                { sLine = "FILENAME_STN2 = "; sw.WriteLine(sLine); }
                rdS.Close();
                cmdS.Dispose();

                //  pro Praporek
                string sSQL_x = ("SELECT batch_no,doc_class,MAX(file_name) AS FileName FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%PRO%'" + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));

                // nechci to brát z c_labels_ad, tak to hledám přímo
                string sSQL_P = "SELECT (c.doc_class || '-' || c.doc_no || '-' ||c.doc_rev || '.' || c.format_size) as Nalepka" +
                        " FROM IFSAPP.DOC_REFERENCE_OBJECT a, IFSAPP.DOC_STRUCTURE b,IFSAPP.DOC_ISSUE_REFERENCE c" +
                        " WHERE a.doc_class = b.doc_class" +
                        " AND a.doc_no = b.doc_no" +
                        " AND a.doc_rev = b.doc_rev" +
                        " AND b.sub_doc_class = c.doc_class" +
                        " AND b.sub_doc_no  = c.doc_no" +
                        " AND b.sub_doc_rev = c.doc_rev" +
                        " AND a.lu_name = 'CustomerOrderLine'" +
                        //" AND a.key_ref LIKE  '%ORDER_NO=" + _Customer_order + "%'" +
                        //" AND a.key_ref LIKE  '%LINE_NO=" + _LINE_NO + "%'" +
                        //" AND a.key_ref LIKE  '%REL_NO=" + _RELEASE_NO + "%'" +
                        " AND a.key_value LIKE  '" + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";


                System.Data.OracleClient.OracleCommand cmdP = new System.Data.OracleClient.OracleCommand(sSQL_P, cn);
                System.Data.OracleClient.OracleDataReader rdP = cmdP.ExecuteReader();
                if (!rdP.HasRows) // NOT !!!!!!!!!!!
                {
                    if ((checkBoxOnlyStockOrders.Checked == true))
                    {
                        sLine = "FILENAME_PRO = ";
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        sLine = "FILENAME_PRO = ";
                        sw.WriteLine(sLine);
                        throw new Exception("Nenašel nálepky PRO !!");
                    }
                }
                else // má řádky
                {
                    // co když tam jsou 2 ZPL ? Ber první
                    int iRows = 0;
                    bool bPRO_BTW = false;
                    while (rdP.Read())
                    {
                        if (rdP[0].GetType() != typeof(DBNull))
                        {
                            sFILENAME_PRO_PRIPONA = Path.GetExtension(rdP.GetString(0).Trim());
                            if (sFILENAME_PRO_PRIPONA == string.Empty)
                            { }
                            else { iRows++; }
                        }
                    }
                    if (iRows == 0)
                    {
                        sLine = "FILENAME_PRO = ";
                        sw.WriteLine(sLine);
                    }
                    if (iRows > 1) // Zvol BTW nebo ZPL 
                    {
                        // co když jsou tam 2 ZPL
                        if (MessageBox.Show("Nalezeny 2 nálepky pro praporek, chcete zařadit do txt praporek BTW ?" +
                            Environment.NewLine + "Pokud ne, vezme se ZPL", "BTW nebo ZPL", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            bPRO_BTW = true;
                        }
                    }
                    rdP = cmdP.ExecuteReader();
                    int iPor = 0;
                    while (rdP.Read())
                    {
                        if (rdP[0].GetType() == typeof(DBNull))
                        {
                            sLine = "FILENAME_PRO = ";
                        }
                        else
                        {
                            string sPRO = rdP.GetString(0).Trim();
                            sFILENAME_PRO_PRIPONA = Path.GetExtension(sPRO);
                            if (sFILENAME_PRO_PRIPONA != string.Empty)
                            {
                                if (iRows == 1) // existuje jen l řádek
                                {
                                    sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                    sFILENAME_PRO = rdP.GetString(0).Trim();
                                    sw.WriteLine(sLine);
                                }
                                else
                                {
                                    if (sPRO.Contains(".BTW") && bPRO_BTW == true && iRows > 1)
                                    {
                                        sLine = ("FILENAME_PRO2 = " + rdP.GetString(0).Trim());
                                        sFILENAME_PRO = rdP.GetString(0).Trim();
                                        sw.WriteLine(sLine);
                                    }
                                    if (sPRO.Contains(".ZPL") && bPRO_BTW == false && iRows > 1)
                                    {
                                        if (iPor==0)
                                        {
                                            sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                            iPor++;                                     
                                        }
                                        else
                                        {
                                            sLine = ("FILENAME_PRO2 = " + rdP.GetString(0).Trim());

                                        }
                                        //sFILENAME_PRO = rdP.GetString(0).Trim();
                                        sw.WriteLine(sLine);
                                    }
                                }
                                iPor++;
                            }
                        }
                    }
                    rdP.Close();
                }
                cmdP.Dispose();

                // DOC_CLASS
                string sDOC_CLASS = string.Empty;
                string sDOC_NO = string.Empty;
                string sDOC_REV = string.Empty;
                sSQL_S = Globals.SQL_DOC_CLASS.Replace("xxORDER_NO", _Customer_order);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", _LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "KTN");
                sErr = ff.Oracle_GetValue3(sConString, sSQL_S, ref sDOC_CLASS, ref sDOC_NO, ref sDOC_REV);

                //CATALOG_NO
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCatalog_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_NO = " + sCatalog_no;
                sw.WriteLine(sLine);

                //PART_NO
                if (_NELA == NelaFile_MachineName.E3PR || _NELA == NelaFile_MachineName.DA001)
                {
                    sSQL_1 = "SELECT part_no FROM IFSAPP.CUSTOMER_ORDER_LINE_TAB WHERE " + sWhereOrderLineKomplet;
                    string sPart_No = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                    sLine = "PART_NO = " + sPart_No;
                    sw.WriteLine(sLine);
                }

                //CATALOG_DESC
                sSQL_1 = "SELECT  CATALOG_DESC FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCATALOG_DESC = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_DESC = " + sCATALOG_DESC;
                sw.WriteLine(sLine);


                //PACKAGE_NUM1 - počáteční číslo kartonu v řádku objednávky
                sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM1 = " + sPACKAGE_NUM1;
                sw.WriteLine(sLine);

                //PACKAGE_NUM2 - konečné číslo kartonu v řádku objednávky
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2 = " + sPACKAGE_NUM2;
                sw.WriteLine(sLine);


                //PACKAGE_QTY (frmCPackage).C_PACKAGE_HEAD.AMOUNT
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    //sSQL_1 = "select IFSAPP.C_Package_Head_Api.Get_Package_Capacity(ORDER_NO, LINE_NO, REL_NO, LINE_ITEM_NO, LU_NAME)" +
                    //    " from IFSAPP.C_PACKAGE_HEAD_CFV" +

                    //    " WHERE " + sWhereOrderKomplet;
                    sSQL_1 = "SELECT IFSAPP.C_Package_Head_Api.Get_Package_Capacity(ORDER_NO, LINE_NO, REL_NO, LINE_ITEM_NO, LU_NAME)" +
                        " FROM IFSAPP.C_PACKAGE_HEAD" +
                        " WHERE " + sWhereOrderKomplet;
                }
                else
                { sSQL_1 = "SELECT  AMOUNT FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet; }
                    
                string sPACKAGE_QTY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iPACKAGE_QTY = 0;
                if (Int32.TryParse(sPACKAGE_QTY, out iPACKAGE_QTY))
                {
                }
                else
                {
                    MessageBox.Show("Nenašel množství v kartonu !!"); return;
                };
                sLine = "PACKAGE_QTY = " + sPACKAGE_QTY;
                sw.WriteLine(sLine);

                //SEQ_NUM -- počet vnitřních obalů
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "select quantity from IFSAPP.C_PACKAGE_HEAD a, IFSAPP.PACKING_INSTRUCTION_NODE b" +
                        " WHERE a.packing_instruction_id =  b.packing_instruction_id AND b.handling_unit_type_id = 'INNER' AND b.node_id = 2 AND " +
                        sWhereOrderKomplet;
                }
                else
                { 
                    sSQL_1 = "SELECT IFSAPP.C_PACKAGE_HEAD_API.Get_Inner_Package_Count(PACKAGE_TYPE,TEMPLATE_ID) FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }
                string sSEQ_NUM = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                if (_NELA == NelaFile_MachineName.E3PR || _NELA == NelaFile_MachineName.DA001)
                {
                    sLine = "SEQ_NUM = " + sSEQ_NUM;
                    sw.WriteLine(sLine);
                }

                //SMALL_PACKAGE_QTY 120/4
                int iSEQ_NUM = 0;
                if (Int32.TryParse(sSEQ_NUM, out iSEQ_NUM)) { } else { MessageBox.Show("Nenašel počet vnitřních obalů !!"); return; };
                int iSMALL_PACKAGE_QTY = iPACKAGE_QTY / iSEQ_NUM;
                sLine = "SMALL_PACKAGE_QTY = " + iSMALL_PACKAGE_QTY.ToString();
                sw.WriteLine(sLine);
                //

                ////PACKAGE_NUM2BEG
                //sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                //string sPACKAGE_NUM2BEG = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //sLine = "PACKAGE_NUM2BEG = " + sPACKAGE_NUM2BEG;
                //sw.WriteLine(sLine);

                ////PACKAGE_NUM2END
                //sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                //string sPACKAGE_NUM2END = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //sLine = "PACKAGE_NUM2END = " + sPACKAGE_NUM2END;
                //sw.WriteLine(sLine);

                ////SEQ_TOTAL - počet skládaček sSMALL_PACKAGE_QTY
                //sLine = "SEQ_TOTAL = " + iSEQ_NUM.ToString();
                //sw.WriteLine(sLine);

                //BOX_PCS
                sLine = "BOX_PCS = " + iSMALL_PACKAGE_QTY.ToString();
                sw.WriteLine(sLine);
                //TOTAL_PCS
                //SEQ_NUM -- počet vnitřních obalů
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT SUM(amount) FROM IFSAPP.C_PACKAGE_DETAIL WHERE package_type = 'INNER' AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT SUM(amount) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                }
                    
                string sTOTAL_PCS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TOTAL_PCS = " + sTOTAL_PCS;
                sw.WriteLine(sLine);

                // u NA nebývá úsporné balení, ve všech kartonech je stejná položka
                //PACKAGE_GROSS_WEIGHT 
                //SEQ_NUM -- počet vnitřních obalů
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }
                string sPACKAGE_GROSS_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_GROSS_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_GROSS_WEIGHT));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma,DecimalSeparator.Dot):sLine);
                sw.WriteLine(sLine);

                //ROUND_NET_WEIGHT
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }
                string sROUND_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ROUND_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##}", Convert.ToDouble(sROUND_NET_WEIGHT));
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                {
                    sw.WriteLine(sLine);
                }

                //PACKAGE_NET_WEIGHT
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }

                string sPACKAGE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:#####.0}", Convert.ToDouble(sPACKAGE_NET_WEIGHT));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sLine);
                sw.WriteLine(sLine);

                //EXPLOSIVE_NET_WEIGHT
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }

                string sEXPLOSIVE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "EXPLOSIVE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:0.000}", Math.Round(Convert.ToDouble(sEXPLOSIVE_NET_WEIGHT),3));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sLine);
                sw.WriteLine(sLine);

                //UN_NO
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  uncode FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  un FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }
                string sUN = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "UN_NO = " + sUN;
                sw.WriteLine(sLine);

                //ADR_RID_CLASS
                sSQL_1 = "SELECT class FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  hazardclass FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE ROWNUM = 1 AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT class FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }

                string sADR_RID_CLASS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ADR_RID_CLASS = " + sADR_RID_CLASS;
                sw.WriteLine(sLine);
                //DETONATOR_CLASS
                string sDETONATOR_CLASS = string.Empty;
                sLine = "DETONATOR_CLASS = "; // u DA není
                                              //sw.WriteLine(sLine);

                //CONSUMPTION_DATE
                sSQL_1 = "SELECT part_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sPart_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT consumption_month FROM IFSAPP.a0000_polozky_char WHERE part_no = '" + sPart_no + "' AND contract = 'AD'";
                string sConsumption_month = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = "SELECT to_char(add_months(promised_delivery_date," + sConsumption_month + "),'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + 
                    sWhereOrderKomplet + " AND (line_item_no = 0 OR line_item_no = -1)";
                string sCONSUMPTION_DATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONSUMPTION_DATE = " + sCONSUMPTION_DATE;
                if (_NELA != NelaFile_MachineName.E3PR && _NELA != NelaFile_MachineName.DA001)
                { 
                    sw.WriteLine(sLine);
                }

                //NOMINAL_TIME_NA = N016
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N016'";
                string sNOMINAL_TIME_NA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_NA = " + sNOMINAL_TIME_NA;

                //NOMINAL_TIME_NA_MULTI = 
                //string sNOMINAL_TIME_NA_MULTI = string.Empty;
                sLine = "NOMINAL_TIME_NA_MULTI = "; // u DA není

                //ST_LENGTH = N005
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N016'";
                string sST_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ST_LENGTH = " + sST_LENGTH; // u DA není

                //DELAY_LEVEL = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E004'";
                string sDELAY_LEVEL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_LEVEL = " + sDELAY_LEVEL; // u DA není
                sw.WriteLine(sLine);

                //DELAY_STEP = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L004'";
                string sDELAY_STEP = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_STEP = " + sDELAY_STEP; // u DA není
                //sw.WriteLine(sLine);

                //DELAY_STEP_MULTI  
                sLine = "DELAY_STEP_MULTI = "; // u DA není

                // CAVITY_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L003'";
                string sCAVITY_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CAVITY_MATERIAL  = " + sCAVITY_MATERIAL;
                sw.WriteLine(sLine);

                //SUPPLY_WIRE_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L008'";
                string sSUPPLY_WIRE_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SUPPLY_WIRE_MATERIAL = " + sSUPPLY_WIRE_MATERIAL;
                sw.WriteLine(sLine);

                //WIRE_ISULATOR_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L009'";
                string sWIRE_ISULATOR_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_ISULATOR_MATERIAL = " + sWIRE_ISULATOR_MATERIAL;
                sw.WriteLine(sLine);

                //WIRE_LENGTH 
                sSQL_1 = "SELECT b.characteristic_value_desc FROM IFSAPP.INV_PART_DISCRETE_CHAR a,IFSAPP.discrete_charac_value b" +
                    " WHERE a.characteristic_code = b.characteristic_code AND a.attr_value = b.characteristic_value" + 
                    " AND  a.part_no = '" + sPart_no + "' AND a.contract = 'AD'" +
                    " AND a.characteristic_code = 'E005'";
                string sWIRE_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_LENGTH = " + sWIRE_LENGTH;
                sw.WriteLine(sLine);

                //NOMINAL_TIME_EA 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sNOMINAL_TIME_EA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_EA = " + sNOMINAL_TIME_EA;
                sLine = (_Preprog ? "NOMINAL_TIME_EA = PREPROG" : sLine);
                sw.WriteLine(sLine);




                //SEQ_NUM_SHIFT 
                string sSEQ_NUM_SHIFT = string.Empty;
                sLine = "SEQ_NUM_SHIFT = "; // u DA není

                //TESTING_STATION_NUMBER
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKompletForSalesPart;
                string sSales_part = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE010'";
                string sTESTING_STATION_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_STATION_NUMBER = " + sTESTING_STATION_NUMBER;
                sw.WriteLine(sLine);

                //TESTING_PROTOCOL_NUMBER
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE011'";
                string sTESTING_PROTOCOL_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_PROTOCOL_NUMBER = " + sTESTING_PROTOCOL_NUMBER;
                sw.WriteLine(sLine);

                //ST_COLOR - N008 (např. žlutá)
                sSQL_1 = "SELECT IFSAPP.discrete_charac_value_api.Get_Characteristic_Value_Desc(a.characteristic_code, attr_value)  FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N008'";
                string sST_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ST_COLOR = " + sST_COLOR;
                //


                //ST_PRODUCER
                sLine = "ST_PRODUCER = " + sPRODUCER;

                //STOPPER = NL03 (YES,NO)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL03'";
                string sSTOPPER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "STOPPER = " + sSTOPPER;

                //CONNECTOR = NL01  (YES,NO)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL01'";
                string sCONNECTOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONNECTOR = " + sCONNECTOR;

                //SEALING = NL02
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL02'";
                string sSEALING = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SEALING = " + sSEALING;

                //DETONATOR_LENGTH = N017 (délka dutinky)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N017'";
                string sDETONATOR_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DETONATOR_LENGTH = " + sDETONATOR_LENGTH;

                //FLAG_COLOR = NL04
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL04'";
                string sFLAG_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "FLAG_COLOR = " + sFLAG_COLOR;

                //TN_COLOR = NL05
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL05'";
                string sTN_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TN_COLOR = " + sTN_COLOR;

                //DATE
                //sSQL_1 = "SELECT to_char(wanted_delivery_date,'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                sSQL_1 = "SELECT to_char(promised_delivery_date,'YYMMDD') FROM IFSAPP.customer_order_line WHERE " +
                        sWhereOrderKomplet + " AND (line_item_no = 0 OR line_item_no = -1)";
                string sDATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DATE = " + sDATE;
                sw.WriteLine(sLine);

                //EXPIRATION
                sLine = "EXPIRATION = " + sCONSUMPTION_DATE;
                sw.WriteLine(sLine);

                //CUSTOMER_AUTHORIZE_DECREE
                sSQL_1 = "SELECT dt.output_type FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + _Customer_order + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sOutput_type = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT dt.note_text FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + _Customer_order + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sNote_text = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                switch (sOutput_type)
                {
                    case "LAB01":
                        sLine = "CUSTOMER_AUTHORIZE_DECREE = " + sNote_text;
                        break;
                    case "LAB02":
                        sLine = "ANOTHER_NOTICE = " + sNote_text;
                        break;
                    case "LAB03":
                        sLine = "NOTE = " + sNote_text;
                        break;
                    default:
                        sLine = "NOTE = " + sNote_text;
                        break;
                }
                if (_NELA == NelaFile_MachineName.E3PR && sLine.Contains("NOTE = ")) { }
                else { sw.WriteLine(sLine); }
                //GENERAL
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GENE'";
                string sGENERAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL = " + sGENERAL;
                sw.WriteLine(sLine);
                //GENERAL_1
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GEN01'";
                string sGENERAL_1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL_1 = " + sGENERAL_1;
                sw.WriteLine(sLine);
                //MANUAL
                if (_NELA == NelaFile_MachineName.E3PR || _NELA == NelaFile_MachineName.DA001)
                {
                    sLine = "MANUAL = ";
                    sw.WriteLine(sLine);
                }



                // 96_DMX_PRODUCER = "90"
                // 86_DMX_PACKAGE_QTY  = "30";
                // 87_DMX_DELAY_LEVEL = "97";
                // 88_DMX_DELAY_STEP = "97"
                // 89_DMX_NOMINAL_TIME_EA = "98";
                // 90_DMX_NOMINAL_TIME_NA = "98";
                // 91_DMX_WIRE_LENGTH = "99"
                // 92_DMX_ST_LENGTH = "99"
                // 93_DMX_DATE = "11"
                // 94_DMX_PACKAGE_NUM2 = "250"
                // 95_DMX_SEQ_NUM = "250"
                //
                string sGS_DMX_PRODUCER = "90";
                string sGS_DMX_PACKAGE_QTY = "30";
                string sGS_DMX_DELAY_LEVEL = "97";
                string sGS_DMX_DELAY_STEP = "97";
                string sGS_DMX_NOMINAL_TIME_EA = "98";
                string sGS_DMX_NOMINAL_TIME_NA = "98";
                string sGS_DMX_WIRE_LENGTH = "99";
                string sGS_DMX_ST_LENGTH = "99";
                string sGS_DMX_DATE = "11";
                string sGS_DMX_PACKAGE_NUM2 = "250";
                string sGS_DMX_SEQ_NUM = "250";
                //
                string sDMX_TEXT_PRODUCER = "DMX_PRODUCER = ";
                string sDMX_TEXT_PACKAGE_QTY = "DMX_PACKAGE_QTY = ";
                string sDMX_TEXT_DELAY_LEVEL = "DMX_DELAY_LEVEL = ";
                string sDMX_TEXT_DELAY_STEP = "DMX_DELAY_STEP = ";
                string sDMX_TEXT_NOMINAL_TIME_EA = "DMX_NOMINAL_TIME_EA = ";
                string sDMX_TEXT_NOMINAL_TIME_NA = "DMX_NOMINAL_TIME_NA = ";
                string sDMX_TEXT_WIRE_LENGTH = "DMX_WIRE_LENGTH = ";
                string sDMX_TEXT_ST_LENGTH = "DMX_ST_LENGTH = ";
                string sDMX_TEXT_DATE = "DMX_DATE = ";
                string sDMX_TEXT_PACKAGE_NUM2 = "DMX_PACKAGE_NUM2 = ";
                string sDMX_TEXT_SEQ_NUM = "DMX_SEQ_NUM = ";

                //DMX_PRODUCER
                sLine = sDMX_TEXT_PRODUCER + sGS_DMX_PRODUCER + sPRODUCER;
                sw.WriteLine(sLine);

                //DMX_PACKAGE_QTY
                sLine = sDMX_TEXT_PACKAGE_QTY + sGS_DMX_PACKAGE_QTY + sPACKAGE_QTY;
                sw.WriteLine(sLine);
                //

                //DMX_DELAY_LEVEL
                sLine = sDMX_TEXT_DELAY_LEVEL + sGS_DMX_DELAY_LEVEL + sDELAY_LEVEL;
                sw.WriteLine(sLine);

                //DMX_DEALY_STEP ;
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sDMX_DEALY_STEP_Value = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = sDMX_TEXT_DELAY_STEP + sGS_DMX_DELAY_STEP + sDMX_DEALY_STEP_Value;
                //sw.WriteLine(sLine);

                //DMX_NOMINAL_TIME_EA
                sLine = sDMX_TEXT_NOMINAL_TIME_EA + sGS_DMX_NOMINAL_TIME_EA + sNOMINAL_TIME_EA;
                sw.WriteLine(sLine);

                //DMX_NOMINAL_TIME_NA
                sLine = sDMX_TEXT_NOMINAL_TIME_NA + sGS_DMX_NOMINAL_TIME_NA + sNOMINAL_TIME_NA;
                //sw.WriteLine(sLine); u DA není

                //DMX_ST_LENGTH
                sLine = sDMX_TEXT_ST_LENGTH + sGS_DMX_ST_LENGTH + sST_LENGTH;
                //sw.WriteLine(sLine); ; u DA není

                //DMX_DATE
                sLine = sDMX_TEXT_DATE +  sGS_DMX_DATE + sDATE;
                sw.WriteLine(sLine);

                //DMX_PACKAGE_NUM2
                sLine =  sDMX_TEXT_PACKAGE_NUM2 + sGS_DMX_PACKAGE_NUM2+  sPACKAGE_NUM2;
                sw.WriteLine(sLine);

                //DMX_SEQ_NUM
                string sDMX_SEQ_NUM = iSEQ_NUM.ToString();
                sLine = sDMX_TEXT_SEQ_NUM + sGS_DMX_SEQ_NUM + sDMX_SEQ_NUM;
                sw.WriteLine(sLine);

                //DMX_WIRE_LENGTH
                string sDMX_WIRE_LENGTH = sWIRE_LENGTH;
                sLine = sDMX_TEXT_WIRE_LENGTH + sGS_DMX_WIRE_LENGTH + sDMX_WIRE_LENGTH;
                sw.WriteLine(sLine);

                //AUTO_LABEL - nové aplikační identifikátory (je to v SL), podle nich si TRC vybírá praporek
                sw.WriteLine((sAUTO_LABEL.Contains("AUTO_LABEL =")? sAUTO_LABEL : "AUTO_LABEL =" + sAUTO_LABEL));
                //
                // UFI_CODE
                string sSQL_UFI = "select ch.CF$_UFI_CODE_A1149 from IFSAPP.A1135_INVPART_CHAR_ALL_CFV ch where ch.CONTRACT = 'AD' and ch.PART_NO = '" + _PART_NO + "' AND ch.CHARACTERISTIC_CODE = 'W001'";
                string sUFI_CODE = ff.Oracle_GetValue(sConString, sSQL_UFI, ref sErr);
                sLine = "UFI_CODE = " + sUFI_CODE;
                sw.WriteLine(sLine);
                //
                if (checkBoxDETONATOR_APLIC.Checked) // doplněno 22.9.2014 Senkerik
                {
                    sLine = "DETONATOR_APPLIC = NO";
                    sw.WriteLine(sLine);
                }
                // datatable
                int iBeg = 0;
                int iEnd = 0;
                Int32.TryParse(sPACKAGE_NUM1, out iBeg);
                Int32.TryParse(sPACKAGE_NUM2, out iEnd);

                for (int ii = iBeg; ii <= iEnd; ii++)
                {
                    DataRow dr = _dtNalepky.NewRow();
                    dr[btw.BATCH_NO] = sBatchNo;
                    dr[btw.ADR_RID_CLASS] = sADR_RID_CLASS;
                    dr[btw.AUTO_LABEL] = sAUTO_LABEL;
                    dr[btw.BOX_PCS] = iSEQ_NUM.ToString();
                    dr[btw.BRIDGE_RESIST] = string.Empty;
                    dr[btw.CATALOG_DESC] = sCATALOG_DESC;
                    dr[btw.CATALOG_NO] = sCatalog_no;
                    dr[btw.CAVITY_MATERIAL] = string.Empty;
                    dr[btw.CONSUMPTION_DATE] = sCONSUMPTION_DATE;
                    dr[btw.CUSTOMER_AUTHORIZE_DECREE] = string.Empty;
                    dr[btw.CUSTOMER_PO_NO] = sCUSTOMER_PO_NO;
                    dr[btw.DATE] = sDATE;
                    dr[btw.DATE_CODE] = string.Empty;
                    dr[btw.DELAY_FRACTION] = string.Empty;
                    dr[btw.DELAY_LEVEL] = string.Empty;
                    dr[btw.DELAY_STEP] = sDELAY_STEP;
                    dr[btw.DETONATOR_CLASS] = sDETONATOR_CLASS;
                    dr[btw.DMX_DATE] = sGS_DMX_DATE;
                    dr[btw.DMX_DELAY_LEVEL] = string.Empty;
                    dr[btw.DMX_DELAY_STEP] = sDELAY_STEP;
                    dr[btw.DMX_NOMINAL_TIME_EA] = string.Empty;
                    dr[btw.DMX_NOMINAL_TIME_NA] = string.Empty;
                    dr[btw.DMX_PACKAGE_NUM2] = string.Empty;
                    dr[btw.DMX_PACKAGE_QTY] = string.Empty;
                    dr[btw.DMX_PRODUCER] = sPRODUCER;
                    dr[btw.DMX_SEQ_NUM] = string.Empty;
                    dr[btw.DMX_ST_LENGTH] = string.Empty;
                    dr[btw.DMX_WIRE_LENGTH] = string.Empty;
                    dr[btw.DOC_CLASS] = sDOC_CLASS;
                    dr[btw.DOC_NO] = sDOC_NO;
                    dr[btw.DOC_REV] = sDOC_REV;
                    dr[btw.DUPLICATE] = string.Empty;
                    dr[btw.EXPIRATION] = sCONSUMPTION_DATE;
                    dr[btw.EXPLOSIVE_NET_WEIGHT] = sEXPLOSIVE_NET_WEIGHT;
                    dr[btw.FILE_NAME] = sFILENAME_KTN;
                    dr[btw.GENERAL] = sGENERAL;
                    dr[btw.GENERAL_1] = sGENERAL_1;
                    dr[btw.IGNITION_CUE_J] = string.Empty;
                    dr[btw.IGNITION_CUE_MJ] = string.Empty;
                    dr[btw.PACKAGE_GROSS_WEIGHT] = sPACKAGE_GROSS_WEIGHT;
                    dr[btw.PACKAGE_NET_WEIGHT] = sPACKAGE_NET_WEIGHT;
                    dr[btw.PACKAGE_NUM1] = sPACKAGE_NUM1;
                    dr[btw.PACKAGE_NUM2] = sPACKAGE_NUM2;
                    dr[btw.PACKAGE_QTY] = sPACKAGE_QTY;
                    dr[btw.PIECE_NUM] = string.Empty;
                    dr[btw.PRODUCER] = sPRODUCER;
                    dr[btw.RECEIVER_ADDRESS1] = sReceiverAddress1;
                    dr[btw.RECEIVER_ADDRESS2] = sReceiverAddress2;
                    dr[btw.RECEIVER_CITY] = sReceiverCity;
                    dr[btw.RECEIVER_COMPANY_NAME2] = sReceiverCompanyName;
                    dr[btw.RECEIVER_COUNTRY] = sReceiverCountry;
                    dr[btw.RECEIVER_EMAIL] = sReceiverContactEmail;
                    dr[btw.RECEIVER_FAX] = sReceiverContactFax;
                    dr[btw.RECEIVER_MOBILE] = sReceiverContactMobile;
                    dr[btw.RECEIVER_NAME] = sReceiverContactName;
                    dr[btw.RECEIVER_PHONE] = sReceiverContactPhone;
                    dr[btw.RECEIVER_ZIP_CODE] = sReceiverZIP_CODE;
                    dr[btw.RELEASE_NO] = _RELEASE_NO;
                    dr[btw.ROUND_NET_WEIGHT] = sROUND_NET_WEIGHT;
                    dr[btw.SAFETY_CUE_J] = string.Empty;
                    dr[btw.SAFETY_STREAM] = string.Empty;
                    dr[btw.SEQ_NUM] = sSEQ_NUM; // počet skládaček v obalu
                    dr[btw.SMALL_PACKAGE_QTY] = iSMALL_PACKAGE_QTY.ToString();
                    dr[btw.STATION_NUM] = sTESTING_STATION_NUMBER ;
                    dr[btw.ST_LENGTH] = sST_LENGTH;
                    dr[btw.SUPPLY_WIRE_MATERIAL] = sSUPPLY_WIRE_MATERIAL;
                    dr[btw.TESTING_PROTOCOL_NUMBER] = sTESTING_PROTOCOL_NUMBER;
                    dr[btw.TESTING_STATION_NUMBER] = sTESTING_STATION_NUMBER;
                    dr[btw.TITLE] = sTitle;
                    dr[btw.TOTAL_PACKAGE_QTY] = sPACKAGE_QTY;
                    dr[btw.TOTAL_RESIST] = string.Empty;
                    dr[btw.UNIQUE_PACKAGE_ID] = string.Empty;
                    dr[btw.UNIQUE_SUBPACKAGE_ID] = string.Empty;
                    dr[btw.UN_NO] = sUN;
                    dr[btw.VENDOR_ADDRESS_ID] = string.Empty;
                    dr[btw.VENDOR_CITY] = string.Empty;
                    dr[btw.VENDOR_COMPANY] = string.Empty;
                    dr[btw.VENDOR_COUNTRY] = string.Empty;
                    dr[btw.VENDOR_NAME] = string.Empty;
                    dr[btw.VENDOR_STREET] = string.Empty;
                    dr[btw.VENDOR_ZIP_CODE] = string.Empty;
                    dr[btw.WIRE_ISULATOR_MATERIAL] = string.Empty;
                    dr[btw.WIRE_LENGTH] = string.Empty;
                    dr[btw.WIRE_LENGTH_TRACE] = string.Empty;
                    dr[btw.WORKER1] = string.Empty;
                    dr[btw.WORKER2] = string.Empty;


                    _dtNalepky.Rows.Add(dr);
                    for (int jj = 1; jj <= iSEQ_NUM; jj++)
                    {
                    }
                }
                // nové 
                //AUTO_LABEL =
                //RFID_TAG = Y
                //PRE_PROGRAM_DET = Y
                //R_FILTER = EIM
                //R_EIM_VERSION = 0001
                //R_ASIC_VER = 2.11
                //R_LOCK = 0000
                //CUSTOM_ORDER_LINE; ORDINAL_NUMBER_RANGE; DELAY; DELAY_TXT; BOX_NUMBER_RANGE

                string sRFID_TAG ="Y";
                sLine = "RFID_TAG = " + sRFID_TAG;
                sw.WriteLine(sLine);
                //
                if (_Preprog)
                { sLine = "PRE_PROGRAM_DET = Y" ; } else { sLine = "PRE_PROGRAM_DET = N"; }
                sw.WriteLine(sLine);
                //
                sLine = "R_FILTER = EIM";
                sw.WriteLine(sLine);
                //
                sLine = "R_EIM_VERSION = 0001";
                sw.WriteLine(sLine);
                //
                sLine = "R_ASIC_VER = 2.11";
                sw.WriteLine(sLine);
                //
                sLine = "R_LOCK = 0000";
                sw.WriteLine(sLine);
                //


                // preprog
                if (_Preprog)
                {
                    if (_NELA == NelaFile_MachineName.E3PR) // po staru
                    {
                        // hlavičku z txt souboru vynechám a dám svoji
                        sLine = "CUSTOM_ORDER_LINE;ORDINAL_NUMBER_RANGE;DELAY;DELAY_TXT;BOX_NUMBER_RANGE";
                        sw.WriteLine(sLine);
                        // bere se z db
                        string sRadek = string.Empty;
                        string sSQL_Prog = "SELECT a.BATCH_NO,a.item_from,a.item_to,a.delay_ms,a.delay_no,a.box_from,a.box_to" +
                            " FROM IFSAPP.A1135_PROG a WHERE batch_no = '" + _ZO_BatchNo + "'" +
                            " ORDER BY a.item_from";
                        System.Data.OracleClient.OracleCommand cmdProg = new System.Data.OracleClient.OracleCommand(sSQL_Prog, cn);
                        System.Data.OracleClient.OracleDataReader rdProg = cmdProg.ExecuteReader();
                        if (!rdProg.HasRows) // NOT !!!!!!!!!!!
                        {
                            MessageBox.Show("Zák. objednávka nemá zadané programování !!" + Environment.NewLine + _ZO_BatchNo);
                            return;
                        }
                        else // má řádky
                        {
                            while (rdProg.Read())
                            {
                                if (rdProg[0].GetType() != typeof(DBNull))
                                {
                                    sRadek = rdProg["BATCH_NO"].ToString() + ";" + 
                                        rdProg["item_from"].ToString() + ".." + rdProg["item_to"].ToString() + ";" + 
                                        rdProg["delay_ms"].ToString() + ";" + rdProg["delay_no"].ToString() + ";" + 
                                        rdProg["box_from"].ToString() + ".." + rdProg["box_to"].ToString();
                                        sw.WriteLine(sRadek); //sBatchNo je již v TXT
                                }
                            }
                        }
                        //else
                        //{ 
                        //    int ii = 0;
                        //    string[] lines = System.IO.File.ReadAllLines(sFilePreprog);
                        //    foreach (string line in lines)
                        //    {
                        //        ii++;
                        //        if (ii > 1)
                        //        {
                        //            //sw.WriteLine(sBatchNo + ";" + line);
                        //            sw.WriteLine(line); //sBatchNo je již v TXT
                        //        }
                        //    }
                        //}
                    }
                    if (_NELA == NelaFile_MachineName.Elektra)
                    {
                        // ELEKTRA 
                        // taky bude z DB
                        //sLine = "CUSTOM_ORDER_LINE; ORDINAL_NUMBER_RANGE; DELAY; DELAY_TXT; BOX_NUMBER_RANGE";
                        sLine = "CUSTOM_ORDER_LINE;UNIT_FROM;UNIT_TO;DELAY;DELAY_NO;BOX_FROM;BOX_TO";
                        sw.WriteLine(sLine);
                        int ii = 0;
                        string[] lines = System.IO.File.ReadAllLines(sFilePreprog);
                        foreach (string line in lines)
                        {
                            ii++;
                            if (ii>1)
                            {
                                sw.WriteLine(sBatchNo + ";" +line);
                            }
                        }
                    }
                }
                else // neprogramuje se
                {
                    if (_NELA == NelaFile_MachineName.E3PR)
                    {
                        // S p. Machem je dohoda, že rozbušky jsou vždy od 1 pro každý řádek objednávky: 
                        sLine = "CUSTOM_ORDER_LINE;ORDINAL_NUMBER_RANGE;DELAY;DELAY_TXT;BOX_NUMBER_RANGE";
                        sw.WriteLine(sLine);

                        //int iFrom = (int.Parse(sPACKAGE_NUM1)==1?1:iPACKAGE_QTY * (int.Parse(sPACKAGE_NUM1)-1) + 1 );
                        //int iKusu = iPACKAGE_QTY * int.Parse(sPACKAGE_NUM2);
                        // select IFSAPP.CUSTOMER_ORDER_LINE_API.GET_ORDER_COMPONENT_SUM('V95031') FROM dual
                        sSQL_1 = "select IFSAPP.CUSTOMER_ORDER_LINE_API.GET_ORDER_COMPONENT_SUM('" + _Customer_order + "') from DUAL";

                        string sKs_All_v_ZO = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                        sLine = sBatchNo + ";1.." + sKs_All_v_ZO + ";712"; //;????;" + sPACKAGE_NUM1 + ".." + sPACKAGE_NUM2;
                        sw.WriteLine(sLine);

                    }
                }
                //string sTime = String.Format("{0}-{1}-{2} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                string sSQLp = "INSERT INTO nela(ZO_BatchNo,VZ_BatchNo,Cesta,DatumExportu,Uzivatel,Nela) VALUES(" +
                            "'" + _ZO_BatchNo + "'," +
                            "'" + _VZ_BatchNo + "'," +
                            "'" + sFile.Trim() + "'," +
                            "'" + sTime + "','" + _UserName.Trim() + "'," +
                            "'" + _NELA.Trim() + "'" +
                            ")";
                //(String.Format("yyyy-MM-dd HH:mm:ss",DateTime.Now) + ("\'" + ")")))))))))))));
                if (MSSQL_Execute(sSQLp) == false)
                {
                    MessageBox.Show("fce: MSSQL_Execute neproběhla v pořádku.");
                }
                Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
                cn.Close();
                sw.Dispose();
                sw.Close();
                MessageBox.Show("Soubor je vytvořen.");
                KopirujNalepky(true);
            }
            catch (Exception ex)
            {
                if (sw != null)
                {
                    sw.Close();
                }
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }
            this.Cursor = Cursors.Default;
        }

 
        private void ToolStripMenuItem_Nalepky_Click(object sender, EventArgs e)
        {
            Nalepky();
        }
        private void Nalepky()
        {
            if (_dsNalepky == null)
            {
                MessageBox.Show("dataGridViewNalepky je prázdný !");
                return;
            }
            KopirujNalepky(false);
            frmNalepkyTisk ff = new frmNalepkyTisk();
            ff.dsNalepky = _dsNalepky;
            ff.ColumnListLabels = _ColumnListLabels;
            ff.Show();
        }
        //private void Nalepky_Create_List()
        //{
        //}
        private void NaplnComboNelaCislo()
        {
            comboBoxNelaCislo.Items.Clear();
            comboBoxNelaCislo.Items.Add("");
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.Nela1);
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.Nela2);
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.Nela3);
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.Nela4);
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.E3PR);
            comboBoxNelaCislo.Items.Add(NelaFile_MachineName.DA001);

        }
        private void Columns_NaplnList()
        {
            _ColumnListLabels.Add(BtwHdr.c01_DOC_CLASS);
            _ColumnListLabels.Add(BtwHdr.c02_TITLE);
            _ColumnListLabels.Add(BtwHdr.c03_DOC_NO);
            _ColumnListLabels.Add(BtwHdr.c04_DOC_REV);
            _ColumnListLabels.Add(BtwHdr.c05_FILE_NAME);
            _ColumnListLabels.Add(BtwHdr.c06_CUSTOMER_PO_NO);
            _ColumnListLabels.Add(BtwHdr.c07_ORDER_NO);
            _ColumnListLabels.Add(BtwHdr.c08_LINE_NO);
            _ColumnListLabels.Add(BtwHdr.c09_RELEASE_NO);
            _ColumnListLabels.Add(BtwHdr.c10_LINE_ITEM_NO);
            _ColumnListLabels.Add(BtwHdr.c11_PACKAGE_NUM1);
            _ColumnListLabels.Add(BtwHdr.c12_SEQ_NUM);
            _ColumnListLabels.Add(BtwHdr.c13_PACKAGE_NUM2);
            _ColumnListLabels.Add(BtwHdr.c14_CATALOG_NO);
            _ColumnListLabels.Add(BtwHdr.c15_CATALOG_DESC);
            _ColumnListLabels.Add(BtwHdr.c16_VENDOR_COMPANY);
            _ColumnListLabels.Add(BtwHdr.c17_VENDOR_NAME);
            _ColumnListLabels.Add(BtwHdr.c18_VENDOR_ADDRESS_ID);
            _ColumnListLabels.Add(BtwHdr.c19_VENDOR_STREET);
            _ColumnListLabels.Add(BtwHdr.c20_VENDOR_CITY);
            _ColumnListLabels.Add(BtwHdr.c21_VENDOR_ZIP_CODE);
            _ColumnListLabels.Add(BtwHdr.c22_VENDOR_COUNTRY);
            _ColumnListLabels.Add(BtwHdr.c23_RECEIVER_COMPANY_NAME2);
            _ColumnListLabels.Add(BtwHdr.c24_RECEIVER_ADDRESS1);
            _ColumnListLabels.Add(BtwHdr.c25_RECEIVER_ADDRESS2);
            _ColumnListLabels.Add(BtwHdr.c26_RECEIVER_ZIP_CODE);
            _ColumnListLabels.Add(BtwHdr.c27_RECEIVER_CITY);
            _ColumnListLabels.Add(BtwHdr.c28_RECEIVER_COUNTRY);
            _ColumnListLabels.Add(BtwHdr.c29_RECEIVER_NAME);
            _ColumnListLabels.Add(BtwHdr.c30_RECEIVER_PHONE);
            _ColumnListLabels.Add(BtwHdr.c31_RECEIVER_FAX);
            _ColumnListLabels.Add(BtwHdr.c32_RECEIVER_MOBILE);
            _ColumnListLabels.Add(BtwHdr.c33_RECEIVER_EMAIL);
            _ColumnListLabels.Add(BtwHdr.c34_UNIQUE_PACKAGE_ID);
            _ColumnListLabels.Add(BtwHdr.c35_UNIQUE_SUBPACKAGE_ID);
            _ColumnListLabels.Add(BtwHdr.c36_PACKAGE_QTY);
            _ColumnListLabels.Add(BtwHdr.c37_SMALL_PACKAGE_QTY);
            _ColumnListLabels.Add(BtwHdr.c38_PACKAGE_GROSS_WEIGHT);
            _ColumnListLabels.Add(BtwHdr.c39_ROUND_NET_WEIGHT);
            _ColumnListLabels.Add(BtwHdr.c40_PACKAGE_NET_WEIGHT);
            _ColumnListLabels.Add(BtwHdr.c41_EXPLOSIVE_NET_WEIGHT);
            _ColumnListLabels.Add(BtwHdr.c42_UN_NO);
            _ColumnListLabels.Add(BtwHdr.c43_ADR_RID_CLASS);
            _ColumnListLabels.Add(BtwHdr.c44_DETONATOR_CLASS);
            _ColumnListLabels.Add(BtwHdr.c45_DELAY_LEVEL);
            _ColumnListLabels.Add(BtwHdr.c46_CAVITY_MATERIAL);
            _ColumnListLabels.Add(BtwHdr.c47_IGNITION_CUE_MJ);
            _ColumnListLabels.Add(BtwHdr.c48_IGNITION_CUE_J);
            _ColumnListLabels.Add(BtwHdr.c49_SAFETY_STREAM);
            _ColumnListLabels.Add(BtwHdr.c50_SAFETY_CUE_MJ);
            _ColumnListLabels.Add(BtwHdr.c51_SAFETY_CUE_J);
            _ColumnListLabels.Add(BtwHdr.c52_SUPPLY_WIRE_MATERIAL);
            _ColumnListLabels.Add(BtwHdr.c53_WIRE_ISULATOR_MATERIAL);
            _ColumnListLabels.Add(BtwHdr.c54_WIRE_LENGTH);
            _ColumnListLabels.Add(BtwHdr.c55_WIRE_LENGTH_TRACE);
            _ColumnListLabels.Add(BtwHdr.c56_CONSUMPTION_DATE);
            _ColumnListLabels.Add(BtwHdr.c57_NOMINAL_TIME_EA);
            _ColumnListLabels.Add(BtwHdr.c58_NOMINAL_TIME_NA);
            _ColumnListLabels.Add(BtwHdr.c59_ST_LENGTH);
            _ColumnListLabels.Add(BtwHdr.c60_DELAY_STEP);
            _ColumnListLabels.Add(BtwHdr.c61_BATCH_NO);
            _ColumnListLabels.Add(BtwHdr.c62_TESTING_STATION_NUMBER);
            _ColumnListLabels.Add(BtwHdr.c63_TESTING_PROTOCOL_NUMBER);
            _ColumnListLabels.Add(BtwHdr.c64_CUSTOMER_AUTHORIZE_DECREE);
            _ColumnListLabels.Add(BtwHdr.c65_ANOTHER_NOTICE);
            _ColumnListLabels.Add(BtwHdr.c66_NOTE);
            _ColumnListLabels.Add(BtwHdr.c67_OBJVERSION);
            _ColumnListLabels.Add(BtwHdr.c68_OBJID);
            _ColumnListLabels.Add(BtwHdr.c69_DATE);
            _ColumnListLabels.Add(BtwHdr.c70_EXPIRATION);
            _ColumnListLabels.Add(BtwHdr.c71_TOTAL_RESIST);
            _ColumnListLabels.Add(BtwHdr.c72_BRIDGE_RESIST);
            _ColumnListLabels.Add(BtwHdr.c73_DATE_CODE);
            _ColumnListLabels.Add(BtwHdr.c74_LOT_NUMBER);
            _ColumnListLabels.Add(BtwHdr.c75_STATION_NUM);
            _ColumnListLabels.Add(BtwHdr.c76_WORKERS);
            _ColumnListLabels.Add(BtwHdr.c77_WORKER2);
            _ColumnListLabels.Add(BtwHdr.c78_DUPLICATE);
            _ColumnListLabels.Add(BtwHdr.c79_PIECE_NUM);
            _ColumnListLabels.Add(BtwHdr.c80_DELAY_FRACTION);
            _ColumnListLabels.Add(BtwHdr.c81_GENERAL);
            _ColumnListLabels.Add(BtwHdr.c82_GENERAL_1);
            _ColumnListLabels.Add(BtwHdr.c83_BOX_PCS);
            _ColumnListLabels.Add(BtwHdr.c84_MANUAL);
            _ColumnListLabels.Add(BtwHdr.c85_PRODUCER);
            _ColumnListLabels.Add(BtwHdr.c86_DMX_PACKAGE_QTY);
            _ColumnListLabels.Add(BtwHdr.c87_DMX_DELAY_LEVEL);
            _ColumnListLabels.Add(BtwHdr.c88_DMX_DELAY_STEP);
            _ColumnListLabels.Add(BtwHdr.c89_DMX_NOMINAL_TIME_EA);
            _ColumnListLabels.Add(BtwHdr.c90_DMX_NOMINAL_TIME_NA);
            _ColumnListLabels.Add(BtwHdr.c91_DMX_WIRE_LENGTH);
            _ColumnListLabels.Add(BtwHdr.c92_DMX_ST_LENGTH);
            _ColumnListLabels.Add(BtwHdr.c93_DMX_DATE);
            _ColumnListLabels.Add(BtwHdr.c94_DMX_PACKAGE_NUM2);
            _ColumnListLabels.Add(BtwHdr.c95_DMX_SEQ_NUM);
            _ColumnListLabels.Add(BtwHdr.c96_DMX_PRODUCER);
            _ColumnListLabels.Add(BtwHdr.c97_AUTO_LABEL);
            _ColumnListLabels.Add(BtwHdr.c98_TOTAL_PACKAGE_QTY);
            _ColumnListLabels.Add(BtwHdr.c99_UFI_CODE);

        }

        private void buttonZONapln_Click(object sender, EventArgs e)
        {
            Napln_dataGridViewVZ(string.Empty);
        }
        private void VytvorSouborProDA_OldOrdering()
        {
            this.Cursor = Cursors.WaitCursor;
            //
            string sFILENAME_KTN = string.Empty;
            string sFILENAME_KTN2 = string.Empty;
            string sFILENAME_STN = string.Empty;
            string sFILENAME_PRO = string.Empty;
            string sTitle = string.Empty;
            
            //
            MakeDataTable_Nalepky();
            //

            string sConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;"; ;
            System.Data.OracleClient.OracleConnection cn;
            System.Data.OracleClient.OracleCommand cm;
            System.Data.OracleClient.OracleDataReader rd;
            StreamWriter sw;

            string sLine;
            string sWhereOrderKomplet = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'";
            string sWhereOrderLineKomplet = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'" +
                            " AND  line_item_no = " + _LINE_ITEM_NO;

            string sWhereOrderLineKompletForSalesPart = " order_no = '" + _Customer_order + "'" +
                            " AND  line_no = '" + _LINE_NO + "'" +
                            " AND  rel_no = '" + _RELEASE_NO + "'" +
                            " AND  line_item_no <= 0";

            string sBatchNo = _ZO_BatchNo;
            string sPACKAGE_QTY_MULTI = string.Empty;
            string sNOMINAL_TIME_NA_MULTI = string.Empty;
            string sDelay_step_multi = string.Empty;
            string sAUTO_LABEL = string.Empty;

            string sFile = (textBoxFolderToSave.Text.Trim() + ("\\AD_" + (_VZ_BatchNo + ".TXT")));
            sw = new StreamWriter(sFile, false, System.Text.Encoding.Default);
            try
            {
                int i = 0;
                string sErr = string.Empty;
                string sSQL_1 = string.Empty;
                Funkce ff = new Funkce();
                // 
                cn = new System.Data.OracleClient.OracleConnection(sConString);
                cn.Open();
                // adresy vendor a receiver
                // CUSTOMER_INFO_ADDRESS.ADDRESS2
                sSQL_1 = Globals.SQL_GetIDZakaznika.Replace("xxORDER_NO", _Customer_order);
                string sCustomer_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = Globals.SQL_GetIDAddress.Replace("xxORDER_NO", _Customer_order);
                string sIDAddress = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                //public static string SQL_GetAddress = "SELECT xxColumn  FROM IFSAPP.CUSTOMER_INFO_ADDRESS a where a.customer_id = 'xxCustomer_id' AND a.address_id = 'xxAddress_id'";
                sSQL_1 = Globals.SQL_GetAddress.Replace("xxColumn", "a.ADDRESS1");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);

                string sReceiverAddress1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.ADDRESS1", "a.ADDRESS2");
                string sReceiverAddress2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.ADDRESS2", "a.CITY");
                string sReceiverCity = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.CITY", "a.COUNTRY");
                string sReceiverCountry = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = sSQL_1.Replace("a.COUNTRY", "a.ZIP_CODE");
                string sReceiverZIP_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCompanyNameZakaznika.Replace("xxCustomer_id", sCustomer_no);
                string sReceiverCompanyName = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD.Replace("xxColumn", "a.Name");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactName = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD.Replace("xxColumn", "a.value");
                sSQL_1 = sSQL_1.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactPhone = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_EMAIL.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactEmail = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_FAX.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactFax = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                sSQL_1 = Globals.SQL_GetCustomer_COMM_METHOD_MOBILE.Replace("xxCustomer_id", sCustomer_no);
                sSQL_1 = sSQL_1.Replace("xxAddress_id", sIDAddress);
                string sReceiverContactMobile = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);


                // VENDOR
                //
                sSQL_1 = Globals.SQL_Get_ProducerCode_Prefix;
                string sPRODUCER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_ADDRESS_ID
                sSQL_1 = Globals.SQL_GetVendor_Address_ID;
                string sVENDOR_ADDRESS_ID = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_CITY
                sSQL_1 = Globals.SQL_GetVendor_City;
                string sVENDOR_CITY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COMPANY = 'AD'
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_COMPANY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COUNTRY
                sSQL_1 = Globals.SQL_Get_Language_Code_FromOrder.Replace("xxORDER_NO", _Customer_order);
                string sLanguageCode = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = Globals.SQL_GetVendor_Country.Replace("xxLANGUAGE_CODE", sLanguageCode);
                string sVENDOR_COUNTRY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_NAME
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_NAME = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_STREET

                //
                //btw.VENDOR_ZIP_CODE



                // jedná se o úsporné balení ?
                string sSQL_UsporneBaleni = "SELECT a.order_no FROM  IFSAPP.C_PACKAGE_HEAD a" +
                            " WHERE a.unique_pkg = 'FALSE'" + // lišky z IFS to daly opačně !!! když je false tak se jedná o úsporné balení
                            " AND  a.order_no = '" + _Customer_order + "'" +
                            " AND  a.line_no = '" + _LINE_NO + "'" +
                            " AND  a.rel_no = '" + _RELEASE_NO + "'";

                bool bUsporneBaleni = ff.Oracle_Exists(sConString, sSQL_UsporneBaleni, ref sErr);
                if (bUsporneBaleni)
                {
                    string sSQL_Multi = "SELECT * FROM (SELECT distinct order_no,package_num2,part_no," +
                        "replace(wm_concat(amount) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') PACKAGE_QTY_MULTI," +
                        "replace(wm_concat(nominal_time) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') NOMINAL_TIME_NA_MULTI," +
                        "replace(wm_concat(delay_step_multi) OVER (PARTITION BY package_num2 ORDER BY order_no,package_num2) ,',','+') Delay_step_multi" +
                        " FROM" +
                        "( " +
                        "   SELECT a.order_no,b.package_num2,b.ch_line_item_no,b.amount,b.part_no," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'N016') as nominal_time," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'L004') as delay_step_multi" +
                        "   FROM  IFSAPP.C_PACKAGE_HEAD a, IFSAPP.C_PACKAGE_DETAIL b " +
                        "   WHERE a.order_no = b.order_no" +
                        "   AND a.line_no = b.line_no" +
                        "   AND a.rel_no = b.rel_no" +
                        "   AND a.line_item_no = b.line_item_no" +
                        "   AND a.unique_pkg = 'FALSE'" +
                        "   AND  a.order_no = '" + _Customer_order + "'" +
                        "   AND  a.line_no = '" + _LINE_NO + "'" +
                        "   AND  a.rel_no = '" + _RELEASE_NO + "'" +
                        "   AND  a.line_item_no = '" + _LINE_ITEM_NO + "'" +
                        "   ORDER BY order_no,ch_line_item_no" +
                        "))" +
                        " WHERE part_no = '" + _PART_NO + "'" +
                        " ORDER BY order_no,package_num2";
                    // při úsporném balení vrátí více řádků (kartonů) pro výrobu jedné položky 
                    sPACKAGE_QTY_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "PACKAGE_QTY_MULTI", ref sErr);
                    sNOMINAL_TIME_NA_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "NOMINAL_TIME_NA_MULTI", ref sErr);
                    sDelay_step_multi = ff.Oracle_GetValue(sConString, sSQL_Multi, "delay_step_multi", ref sErr);


                }

                //
                sLine = "[TABLE]";
                sw.WriteLine(sLine);
                sLine = "MACHINE_NAME = " + _NELA;
                sw.WriteLine(sLine);
                sLine = "STATION_NUM = " + _NELA;
                sw.WriteLine(sLine);
                // BATCH_NO
                sLine = ("BATCH_NO = " + sBatchNo);
                sw.WriteLine(sLine);
                // PREPROG 
                if (_Preprog)
                {
                    sLine = ("PREPROG = " + "PREPROG_" + sBatchNo + ".TXT");
                }
                else
                { sLine = "PREPROG = "; }
                sw.WriteLine(sLine);
                // FILENAME_KTN

                // FILENAME_KTN2
                // FILENAME_STN


                // KARTONY
                //string sSQL_K = ("SELECT distinct file_name FROM IFSAPP.c_labels_ad a" + (" WHERE a.order_no = \'"
                //            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                //            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                //            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                //            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%KTN%'" + " ORDER BY a.file_name"))))))))))))));

                // výběr kartonů  ze zakázek z IFS 
                string sSQL_K = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", _Customer_order);
                sSQL_K = sSQL_K.Replace("xxLINE_NO", _LINE_NO);
                sSQL_K = sSQL_K.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_K = sSQL_K.Replace("xxTypNalepky", "KTN");
                System.Data.OracleClient.OracleCommand cmdK = new System.Data.OracleClient.OracleCommand(sSQL_K, cn);
                System.Data.OracleClient.OracleDataReader rdK = cmdK.ExecuteReader();
                if (!rdK.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-KTN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                        sw.WriteLine(sLine);
                    }
                }
                while (rdK.Read())
                {
                    i++;
                    if ((i == 2))
                    {
                        if (rdK[0].GetType() != typeof(DBNull))
                        {
                            sLine = ("FILENAME_KTN2 = " + rdK.GetString(0).Trim());
                        }
                        else
                        {
                            sLine = "FILENAME_KTN2 = ";
                        }
                    }
                    else if ((rdK[0].GetType() != typeof(DBNull)))
                    {
                        sLine = ("FILENAME_KTN = " + rdK.GetString(0).Trim());
                        sFILENAME_KTN = rdK.GetString(0).Trim();
                    }
                    else
                    {
                        sLine = "FILENAME_KTN = ";
                    }
                    sw.WriteLine(sLine);
                }
                if ((i == 1))
                {
                    sLine = "FILENAME_KTN2 =";
                    sw.WriteLine(sLine);
                }
                rdK.Close();
                cmdK.Dispose();


                //  pro Skladacku

                //string sSQL_S = ("SELECT batch_no,doc_class,MAX(file_name) AS FileName FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                //            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                //            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                //            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                //            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%STN%'" + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));

                string sSQL_S = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", _Customer_order);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", _LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "STN");

                System.Data.OracleClient.OracleCommand cmdS = new System.Data.OracleClient.OracleCommand(sSQL_S, cn);
                System.Data.OracleClient.OracleDataReader rdS = cmdS.ExecuteReader();
                if (!rdS.HasRows)
                {
                    if ((checkBoxOnlyStockOrders.Checked == false))
                    {
                        throw new Exception("Nenašel nálepky N-NA-STN !!");
                    }
                    else
                    {
                        sLine = "FILENAME_STN = ";
                        sw.WriteLine(sLine);
                    }
                }
                sLine = "FILENAME_STN = ";
                while (rdS.Read())
                {
                    if (rdS["FileName"].GetType() == typeof(DBNull))
                    {
                        sLine = "FILENAME_STN = ";
                    }
                    else
                    {
                        sLine = ("FILENAME_STN = " + rdS["FileName"].ToString());
                        sFILENAME_STN = rdS["FileName"].ToString();
                        //AUTO_LABEL
                        string sHodnotaPole = rdS["FileName"].ToString();
                        sHodnotaPole = sHodnotaPole.Replace("-A1.BTW", "");
                        sAUTO_LABEL = MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'");
                        sAUTO_LABEL = "AUTO_LABEL = " + sAUTO_LABEL;

                    }
                    sw.WriteLine(sLine);
                    if (rdS["Title"].GetType() != typeof(DBNull))
                    {
                        sTitle = rdS["Title"].ToString();
                    }
                }
                rdS.Close();
                cmdS.Dispose();

                //  pro Praporek
                string sSQL_x = ("SELECT batch_no,doc_class,MAX(file_name) AS FileName FROM IFSAPP.c_labels_ad a " + (" WHERE a.order_no = \'"
                            + (_Customer_order + ("\'" + (" AND  a.line_no = \'"
                            + (_LINE_NO + ("\'" + (" AND  a.release_no = \'"
                            + (_RELEASE_NO + ("\'" + (" AND  a.line_item_no = \'"
                            + (_LINE_ITEM_NO + ("\'" + (" AND doc_class LIKE '%PRO%'" + " GROUP BY batch_no,doc_class ORDER BY batch_no,doc_class"))))))))))))));

                // nechci to brát z c_labels_ad, tak to hledám přímo
                string sSQL_P = "SELECT (c.doc_class || '-' || c.doc_no || '-' ||c.doc_rev || '.' || c.format_size) as Nalepka" +
                        " FROM IFSAPP.DOC_REFERENCE_OBJECT a, IFSAPP.DOC_STRUCTURE b,IFSAPP.DOC_ISSUE_REFERENCE c" +
                        " WHERE a.doc_class = b.doc_class" +
                        " AND a.doc_no = b.doc_no" +
                        " AND a.doc_rev = b.doc_rev" +
                        " AND b.sub_doc_class = c.doc_class" +
                        " AND b.sub_doc_no  = c.doc_no" +
                        " AND b.sub_doc_rev = c.doc_rev" +
                        " AND a.lu_name = 'CustomerOrderLine'" +
                        //" AND a.key_ref LIKE  '%ORDER_NO=" + _Customer_order + "%'" +
                        //" AND a.key_ref LIKE  '%LINE_NO=" + _LINE_NO + "%'" +
                        //" AND a.key_ref LIKE  '%REL_NO=" + _RELEASE_NO + "%'" +
                        " AND a.key_value LIKE  '" + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";


                System.Data.OracleClient.OracleCommand cmdP = new System.Data.OracleClient.OracleCommand(sSQL_P, cn);
                System.Data.OracleClient.OracleDataReader rdP = cmdP.ExecuteReader();
                if (!rdP.HasRows) // NOT !!!!!!!!!!!
                {
                    if ((checkBoxOnlyStockOrders.Checked == true))
                    {
                        sLine = "FILENAME_PRO = ";
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        throw new Exception("Nenašel nálepky PRO !!");
                    }
                }
                else // má řádky
                {
                    int iRows = 0;
                    bool bPRO_BTW = false;
                    while (rdP.Read())
                    {
                        iRows++;
                    }
                    if (iRows > 1) // Zvol BTW nebo ZPL 
                    {
                        if (MessageBox.Show("Nalezeny 2 nálepky pro praporek, chcete zařadit do txt praporek BTW ?" +
                            Environment.NewLine + "Pokud ne, vezme se ZPL", "BTW nebo ZPL", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            bPRO_BTW = true;
                        }
                    }
                    rdP = cmdP.ExecuteReader();
                    while (rdP.Read())
                    {
                        if (rdP[0].GetType() == typeof(DBNull))
                        {
                            sLine = "FILENAME_PRO = ";
                        }
                        else
                        {
                            string sPRO = rdP.GetString(0).Trim();
                            if (iRows == 1) // existuje jen l řádek
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
                                sw.WriteLine(sLine);
                            }
                            if (sPRO.Contains(".BTW") && bPRO_BTW == true && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
                                sw.WriteLine(sLine);
                            }
                            if (sPRO.Contains(".ZPL") && bPRO_BTW == false && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
                                sw.WriteLine(sLine);
                            }
                        }
                    }
                    rdP.Close();
                }
                cmdP.Dispose();
                //ORDER_NO
                sLine = "ORDER_NO = " + _Customer_order;
                sw.WriteLine(sLine);
                //LINE_NO
                if (_NELA != NelaFile_MachineName.E3PR)
                {
                    sLine = "LINE_NO = " + _LINE_NO;
                    sw.WriteLine(sLine);
                }

                //RELEASE_NO
                if (_NELA != NelaFile_MachineName.E3PR)
                {
                    sLine = "RELEASE_NO = " + _RELEASE_NO;
                    sw.WriteLine(sLine);
                }


                //LINE_ITEM_NO
                if (_NELA != NelaFile_MachineName.E3PR)
                {
                    sLine = "LINE_ITEM_NO = " + _LINE_ITEM_NO;
                    sw.WriteLine(sLine);
                }


                //PACKAGE_NUM1 - počáteční číslo kartonu v řádku objednávky
                sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM1 = " + sPACKAGE_NUM1;
                sw.WriteLine(sLine);
                //SEQ_NUM
                sSQL_1 = "SELECT IFSAPP.C_PACKAGE_HEAD_API.Get_Inner_Package_Count(PACKAGE_TYPE,TEMPLATE_ID) FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                string sSEQ_NUM = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                if (_NELA == NelaFile_MachineName.E3PR)
                {
                    sLine = "SEQ_NUM = " + sSEQ_NUM;
                    sw.WriteLine(sLine);
                }
                //PACKAGE_NUM2 - konečné číslo kartonu v řádku objednávky
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2 = " + sPACKAGE_NUM2;
                sw.WriteLine(sLine);
                //CATALOG_NO
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCatalog_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_NO = " + sCatalog_no;
                sw.WriteLine(sLine);

                //PART_NO
                if (_NELA == NelaFile_MachineName.E3PR)
                {
                    sSQL_1 = "SELECT part_no FROM IFSAPP.CUSTOMER_ORDER_LINE_TAB WHERE " + sWhereOrderLineKomplet;
                    string sPart_No = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                    sLine = "PART_NO = " + sPart_No;
                    sw.WriteLine(sLine);
                }

                //CATALOG_DESC
                sSQL_1 = "SELECT  CATALOG_DESC FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCATALOG_DESC = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_DESC = " + sCATALOG_DESC;
                sw.WriteLine(sLine);

                //PACKAGE_QTY (frmCPackage).C_PACKAGE_HEAD.AMOUNT
                sSQL_1 = "SELECT  AMOUNT FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                string sPACKAGE_QTY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iPACKAGE_QTY = 0;
                if (Int32.TryParse(sPACKAGE_QTY, out iPACKAGE_QTY)) { } else { MessageBox.Show("Nenašel množství v kartonu !!"); return; };
                sLine = "PACKAGE_QTY = " + sPACKAGE_QTY;
                sw.WriteLine(sLine);

                //SMALL_PACKAGE_QTY 120/4
                int iSEQ_NUM = 0;
                if (Int32.TryParse(sSEQ_NUM, out iSEQ_NUM)) { } else { MessageBox.Show("Nenašel počet vnitřních obalů !!"); return; };
                int iSMALL_PACKAGE_QTY = iPACKAGE_QTY / iSEQ_NUM;
                sLine = "SMALL_PACKAGE_QTY = " + iSMALL_PACKAGE_QTY.ToString();
                sw.WriteLine(sLine);

                //PACKAGE_GROSS_WEIGHT
                sSQL_1 = "SELECT  package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sPACKAGE_GROSS_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_GROSS_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_GROSS_WEIGHT));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sLine);
                sw.WriteLine(sLine);

                //PACKAGE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sPACKAGE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_NET_WEIGHT));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sLine);
                sw.WriteLine(sLine);

                //EXPLOSIVE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sEXPLOSIVE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "EXPLOSIVE_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:0.000}", Math.Round(Convert.ToDouble(sEXPLOSIVE_NET_WEIGHT), 3));
                sLine = (checkBoxDecimalDelititerChange.Checked ? sLine.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sLine);
                sw.WriteLine(sLine);

                //UN_NO
                sSQL_1 = "SELECT  un FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sUN = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "UN_NO = " + sUN;
                sw.WriteLine(sLine);

                //ADR_RID_CLASS
                sSQL_1 = "SELECT class FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sADR_RID_CLASS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ADR_RID_CLASS = " + sADR_RID_CLASS;
                sw.WriteLine(sLine);

                //
                sSQL_1 = "SELECT part_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sPart_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                //DELAY_LEVEL = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E004'";
                string sDELAY_LEVEL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_LEVEL = " + sDELAY_LEVEL; // u DA není
                sw.WriteLine(sLine);

                // CAVITY_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L003'";
                string sCAVITY_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CAVITY_MATERIAL  = " + sCAVITY_MATERIAL;
                sw.WriteLine(sLine);

                // SUPPLY_WIRE_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L008'";
                string sSUPPLY_WIRE_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SUPPLY_WIRE_MATERIAL = " + sSUPPLY_WIRE_MATERIAL;
                sw.WriteLine(sLine);

                // WIRE_ISULATOR_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L009'";
                string sWIRE_ISULATOR_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_ISULATOR_MATERIAL = " + sWIRE_ISULATOR_MATERIAL;
                sw.WriteLine(sLine);

                // WIRE_LENGTH 
                sSQL_1 = "SELECT b.characteristic_value_desc FROM IFSAPP.INV_PART_DISCRETE_CHAR a,IFSAPP.discrete_charac_value b" +
                    " WHERE a.characteristic_code = b.characteristic_code AND a.attr_value = b.characteristic_value" +
                    " AND  a.part_no = '" + sPart_no + "' AND a.contract = 'AD'" +
                    " AND a.characteristic_code = 'E005'";
                string sWIRE_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_LENGTH = " + sWIRE_LENGTH;
                sw.WriteLine(sLine);

                // NOMINAL_TIME_EA 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sNOMINAL_TIME_EA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_EA = " + sNOMINAL_TIME_EA;
                sLine = (_Preprog ? "NOMINAL_TIME_EA = PREPROG" : sLine);
                sw.WriteLine(sLine);

                // TESTING_STATION_NUMBER
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKompletForSalesPart;
                string sSales_part = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE010'";
                string sTESTING_STATION_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_STATION_NUMBER = " + sTESTING_STATION_NUMBER;
                sw.WriteLine(sLine);

                // TESTING_PROTOCOL_NUMBER
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE011'";
                string sTESTING_PROTOCOL_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_PROTOCOL_NUMBER = " + sTESTING_PROTOCOL_NUMBER;
                sw.WriteLine(sLine);

                // GENERAL
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GENE'";
                string sGENERAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL = " + sGENERAL;
                sw.WriteLine(sLine);

                // GENERAL_1
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GEN01'";
                string sGENERAL_1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL_1 = " + sGENERAL_1;
                sw.WriteLine(sLine);

                // DATE
                sSQL_1 = "SELECT to_char(promised_delivery_date,'YYMMDD') FROM IFSAPP.customer_order_line WHERE " +
                        sWhereOrderKomplet + " AND (line_item_no = 0 OR line_item_no = -1)";
                string sDATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DATE = " + sDATE;
                sw.WriteLine(sLine);

                // EXPIRATION
                sSQL_1 = "SELECT consumption_month FROM IFSAPP.a0000_polozky_char WHERE part_no = '" + sPart_no + "' AND contract = 'AD'";
                string sConsumption_month = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = "SELECT to_char(add_months(promised_delivery_date," + sConsumption_month + "),'YYMMDD') FROM IFSAPP.customer_order_line WHERE " +
                    sWhereOrderKomplet + " AND (line_item_no = 0 OR line_item_no = -1)";
                string sCONSUMPTION_DATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONSUMPTION_DATE = " + sCONSUMPTION_DATE;

                if (_NELA != NelaFile_MachineName.E3PR) //CONSUMPTION_DATE
                {
                    sw.WriteLine(sLine);
                }

                sLine = "EXPIRATION = " + sCONSUMPTION_DATE;
                sw.WriteLine(sLine);

                //PACKAGE_NUM2BEG
                sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2BEG = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2BEG = " + sPACKAGE_NUM2BEG;
                sw.WriteLine(sLine);

                //PACKAGE_NUM2END
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2END = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2END = " + sPACKAGE_NUM2END;
                sw.WriteLine(sLine);

                //SEQ_TOTAL - počet skládaček sSMALL_PACKAGE_QTY
                sLine = "SEQ_TOTAL = " + iSEQ_NUM.ToString();
                sw.WriteLine(sLine);

                //BOX_PCS
                sLine = "BOX_PCS = " + iSMALL_PACKAGE_QTY.ToString();
                sw.WriteLine(sLine);

                //MANUAL
                if (_NELA == NelaFile_MachineName.E3PR)
                {
                    sLine = "MANUAL = ";
                    sw.WriteLine(sLine);
                }

                //TOTAL_PCS
                sSQL_1 = "SELECT SUM(amount) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sTOTAL_PCS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TOTAL_PCS = " + sTOTAL_PCS;
                sw.WriteLine(sLine);

                //PRODUCER
                sLine = "PRODUCER = " + sPRODUCER;
                sw.WriteLine(sLine);

                //CUSTOMER_PO_NO 
                sSQL_1 = "SELECT CUSTOMER_PO_NO FROM IFSAPP.CUSTOMER_ORDER WHERE order_no = '" + _Customer_order + "'";
                string sCUSTOMER_PO_NO = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CUSTOMER_PO_NO = " + sCUSTOMER_PO_NO;
                if (_NELA != NelaFile_MachineName.E3PR)
                {
                    sw.WriteLine(sLine);
                }

                // DMX
                string sGS_DMX_PRODUCER = "90";
                string sGS_DMX_PACKAGE_QTY = "30";
                string sGS_DMX_DELAY_LEVEL = "97";
                string sGS_DMX_DELAY_STEP = "97";
                string sGS_DMX_NOMINAL_TIME_EA = "98";
                string sGS_DMX_NOMINAL_TIME_NA = "98";
                string sGS_DMX_WIRE_LENGTH = "99";
                string sGS_DMX_ST_LENGTH = "99";
                string sGS_DMX_DATE = "11";
                string sGS_DMX_PACKAGE_NUM2 = "250";
                string sGS_DMX_SEQ_NUM = "250";
                //
                string sDMX_TEXT_PRODUCER = "DMX_PRODUCER = ";
                string sDMX_TEXT_PACKAGE_QTY = "DMX_PACKAGE_QTY = ";
                string sDMX_TEXT_DELAY_LEVEL = "DMX_DELAY_LEVEL = ";
                string sDMX_TEXT_DELAY_STEP = "DMX_DELAY_STEP = ";
                string sDMX_TEXT_NOMINAL_TIME_EA = "DMX_NOMINAL_TIME_EA = ";
                string sDMX_TEXT_NOMINAL_TIME_NA = "DMX_NOMINAL_TIME_NA = ";
                string sDMX_TEXT_WIRE_LENGTH = "DMX_WIRE_LENGTH = ";
                string sDMX_TEXT_ST_LENGTH = "DMX_ST_LENGTH = ";
                string sDMX_TEXT_DATE = "DMX_DATE = ";
                string sDMX_TEXT_PACKAGE_NUM2 = "DMX_PACKAGE_NUM2 = ";
                string sDMX_TEXT_SEQ_NUM = "DMX_SEQ_NUM = ";

                //DMX_PRODUCER
                sLine = sDMX_TEXT_PRODUCER + sGS_DMX_PRODUCER + sPRODUCER;
                sw.WriteLine(sLine);


                //DMX_PACKAGE_QTY
                sLine = sDMX_TEXT_PACKAGE_QTY + sGS_DMX_PACKAGE_QTY + sPACKAGE_QTY;
                sw.WriteLine(sLine);

                //DMX_DELAY_LEVEL
                sLine = sDMX_TEXT_DELAY_LEVEL + sGS_DMX_DELAY_LEVEL + sDELAY_LEVEL;
                sw.WriteLine(sLine);

                //DMX_DEALY_STEP;
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sDMX_DEALY_STEP_Value = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = sDMX_TEXT_DELAY_STEP + sGS_DMX_DELAY_STEP + sDMX_DEALY_STEP_Value;
                //sw.WriteLine(sLine); u DA není

                //DMX_NOMINAL_TIME_EA
                sLine = sDMX_TEXT_NOMINAL_TIME_EA + sGS_DMX_NOMINAL_TIME_EA + sNOMINAL_TIME_EA;
                sw.WriteLine(sLine);

                //DMX_NOMINAL_TIME_NA
                //NOMINAL_TIME_NA = N016
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N016'";
                string sNOMINAL_TIME_NA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_NA = " + sNOMINAL_TIME_NA;
                sLine = sDMX_TEXT_NOMINAL_TIME_NA + sGS_DMX_NOMINAL_TIME_NA + sNOMINAL_TIME_NA;
                //sw.WriteLine(sLine); u DA není

                //DMX_ST_LENGTH
                //ST_LENGTH = N005
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N016'";
                string sST_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ST_LENGTH = " + sST_LENGTH; // u DA není
                sLine = sDMX_TEXT_ST_LENGTH + sGS_DMX_ST_LENGTH + sST_LENGTH;
                //sw.WriteLine(sLine); ; u DA není

                //DMX_DATE
                sLine = sDMX_TEXT_DATE + sGS_DMX_DATE + sDATE;
                sw.WriteLine(sLine);

                //DMX_PACKAGE_NUM2
                sLine = sDMX_TEXT_PACKAGE_NUM2 + sGS_DMX_PACKAGE_NUM2 + sPACKAGE_NUM2;
                sw.WriteLine(sLine);

                //DMX_SEQ_NUM
                string sDMX_SEQ_NUM = iSEQ_NUM.ToString();
                sLine = sDMX_TEXT_SEQ_NUM + sGS_DMX_SEQ_NUM + sDMX_SEQ_NUM;
                sw.WriteLine(sLine);

                //DMX_WIRE_LENGTH
                string sDMX_WIRE_LENGTH = sWIRE_LENGTH;
                sLine = sDMX_TEXT_WIRE_LENGTH + sGS_DMX_WIRE_LENGTH + sDMX_WIRE_LENGTH;
                sw.WriteLine(sLine);

                //AUTO_LABEL - nové aplikační identifikátory (je to v SL), podle nich si TRC vybírá praporek
                sw.WriteLine(sAUTO_LABEL);
                //
                if (checkBoxDETONATOR_APLIC.Checked) // doplněno 22.9.2014 Senkerik
                {
                    sLine = "DETONATOR_APPLIC = NO";
                    sw.WriteLine(sLine);
                }

                // KONEC pro Fidricha


                // DOC_CLASS
                string sDOC_CLASS = string.Empty;
                string sDOC_NO = string.Empty;
                string sDOC_REV = string.Empty;
                sSQL_S = Globals.SQL_DOC_CLASS.Replace("xxORDER_NO", _Customer_order);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", _LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", _RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "KTN");
                sErr = ff.Oracle_GetValue3(sConString, sSQL_S, ref sDOC_CLASS, ref sDOC_NO, ref sDOC_REV);

                //ROUND_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sROUND_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ROUND_NET_WEIGHT = " + String.Format(CultureInfo.CurrentCulture, "{0:##}", Convert.ToDouble(sROUND_NET_WEIGHT));
                if (_NELA != NelaFile_MachineName.E3PR)
                {
                    sw.WriteLine(sLine);
                }

                //DETONATOR_CLASS
                string sDETONATOR_CLASS = string.Empty;
                sLine = "DETONATOR_CLASS = "; // u DA není
                                              //sw.WriteLine(sLine);



                //NOMINAL_TIME_NA_MULTI = 
                //string sNOMINAL_TIME_NA_MULTI = string.Empty;
                sLine = "NOMINAL_TIME_NA_MULTI = "; // u DA není



                //DELAY_STEP = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L004'";
                string sDELAY_STEP = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_STEP = " + sDELAY_STEP; // u DA není
                //sw.WriteLine(sLine);

                //DELAY_STEP_MULTI  
                sLine = "DELAY_STEP_MULTI = "; // u DA není

                //SEQ_NUM_SHIFT 
                string sSEQ_NUM_SHIFT = string.Empty;
                sLine = "SEQ_NUM_SHIFT = "; // u DA není

                //ST_COLOR - N008 (např. žlutá)
                sSQL_1 = "SELECT IFSAPP.discrete_charac_value_api.Get_Characteristic_Value_Desc(a.characteristic_code, attr_value)  FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N008'";
                string sST_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ST_COLOR = " + sST_COLOR;
                //


                //ST_PRODUCER
                sLine = "ST_PRODUCER = " + sPRODUCER;

                //STOPPER = NL03 (YES,NO)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL03'";
                string sSTOPPER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "STOPPER = " + sSTOPPER;

                //CONNECTOR = NL01  (YES,NO)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL01'";
                string sCONNECTOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONNECTOR = " + sCONNECTOR;

                //SEALING = NL02
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL02'";
                string sSEALING = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SEALING = " + sSEALING;

                //DETONATOR_LENGTH = N017 (délka dutinky)
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'N017'";
                string sDETONATOR_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DETONATOR_LENGTH = " + sDETONATOR_LENGTH;

                //FLAG_COLOR = NL04
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL04'";
                string sFLAG_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "FLAG_COLOR = " + sFLAG_COLOR;

                //TN_COLOR = NL05
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'NL05'";
                string sTN_COLOR = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TN_COLOR = " + sTN_COLOR;


                //CUSTOMER_AUTHORIZE_DECREE
                sSQL_1 = "SELECT dt.output_type FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + _Customer_order + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sOutput_type = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT dt.note_text FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + _Customer_order + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sNote_text = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                switch (sOutput_type)
                {
                    case "LAB01":
                        sLine = "CUSTOMER_AUTHORIZE_DECREE = " + sNote_text;
                        break;
                    case "LAB02":
                        sLine = "ANOTHER_NOTICE = " + sNote_text;
                        break;
                    case "LAB03":
                        sLine = "NOTE = " + sNote_text;
                        break;
                    default:
                        sLine = "NOTE = " + sNote_text;
                        break;
                }
                if (_NELA == NelaFile_MachineName.E3PR && sLine.Contains("NOTE = ")) { }
                else { sw.WriteLine(sLine); }



                // 96_DMX_PRODUCER = "90"
                // 86_DMX_PACKAGE_QTY  = "30";
                // 87_DMX_DELAY_LEVEL = "97";
                // 88_DMX_DELAY_STEP = "97"
                // 89_DMX_NOMINAL_TIME_EA = "98";
                // 90_DMX_NOMINAL_TIME_NA = "98";
                // 91_DMX_WIRE_LENGTH = "99"
                // 92_DMX_ST_LENGTH = "99"
                // 93_DMX_DATE = "11"
                // 94_DMX_PACKAGE_NUM2 = "250"
                // 95_DMX_SEQ_NUM = "250"
                //
                //

                // datatable
                int iBeg = 0;
                int iEnd = 0;
                Int32.TryParse(sPACKAGE_NUM1, out iBeg);
                Int32.TryParse(sPACKAGE_NUM2, out iEnd);

                for (int ii = iBeg; ii <= iEnd; ii++)
                {
                    DataRow dr = _dtNalepky.NewRow();
                    dr[btw.BATCH_NO] = sBatchNo;
                    dr[btw.ADR_RID_CLASS] = sADR_RID_CLASS;
                    dr[btw.AUTO_LABEL] = sAUTO_LABEL;
                    dr[btw.BOX_PCS] = iSEQ_NUM.ToString();
                    dr[btw.BRIDGE_RESIST] = string.Empty;
                    dr[btw.CATALOG_DESC] = sCATALOG_DESC;
                    dr[btw.CATALOG_NO] = sCatalog_no;
                    dr[btw.CAVITY_MATERIAL] = string.Empty;
                    dr[btw.CONSUMPTION_DATE] = sCONSUMPTION_DATE;
                    dr[btw.CUSTOMER_AUTHORIZE_DECREE] = string.Empty;
                    dr[btw.CUSTOMER_PO_NO] = sCUSTOMER_PO_NO;
                    dr[btw.DATE] = sDATE;
                    dr[btw.DATE_CODE] = string.Empty;
                    dr[btw.DELAY_FRACTION] = string.Empty;
                    dr[btw.DELAY_LEVEL] = string.Empty;
                    dr[btw.DELAY_STEP] = sDELAY_STEP;
                    dr[btw.DETONATOR_CLASS] = sDETONATOR_CLASS;
                    dr[btw.DMX_DATE] = sGS_DMX_DATE;
                    dr[btw.DMX_DELAY_LEVEL] = string.Empty;
                    dr[btw.DMX_DELAY_STEP] = sDELAY_STEP;
                    dr[btw.DMX_NOMINAL_TIME_EA] = string.Empty;
                    dr[btw.DMX_NOMINAL_TIME_NA] = string.Empty;
                    dr[btw.DMX_PACKAGE_NUM2] = string.Empty;
                    dr[btw.DMX_PACKAGE_QTY] = string.Empty;
                    dr[btw.DMX_PRODUCER] = sPRODUCER;
                    dr[btw.DMX_SEQ_NUM] = string.Empty;
                    dr[btw.DMX_ST_LENGTH] = string.Empty;
                    dr[btw.DMX_WIRE_LENGTH] = string.Empty;
                    dr[btw.DOC_CLASS] = sDOC_CLASS;
                    dr[btw.DOC_NO] = sDOC_NO;
                    dr[btw.DOC_REV] = sDOC_REV;
                    dr[btw.DUPLICATE] = string.Empty;
                    dr[btw.EXPIRATION] = sCONSUMPTION_DATE;
                    dr[btw.EXPLOSIVE_NET_WEIGHT] = sEXPLOSIVE_NET_WEIGHT;
                    dr[btw.FILE_NAME] = sFILENAME_KTN;
                    dr[btw.GENERAL] = sGENERAL;
                    dr[btw.GENERAL_1] = sGENERAL_1;
                    dr[btw.IGNITION_CUE_J] = string.Empty;
                    dr[btw.IGNITION_CUE_MJ] = string.Empty;
                    dr[btw.PACKAGE_GROSS_WEIGHT] = sPACKAGE_GROSS_WEIGHT;
                    dr[btw.PACKAGE_NET_WEIGHT] = sPACKAGE_NET_WEIGHT;
                    dr[btw.PACKAGE_NUM1] = sPACKAGE_NUM1;
                    dr[btw.PACKAGE_NUM2] = sPACKAGE_NUM2;
                    dr[btw.PACKAGE_QTY] = sPACKAGE_QTY;
                    dr[btw.PIECE_NUM] = string.Empty;
                    dr[btw.PRODUCER] = sPRODUCER;
                    dr[btw.RECEIVER_ADDRESS1] = sReceiverAddress1;
                    dr[btw.RECEIVER_ADDRESS2] = sReceiverAddress2;
                    dr[btw.RECEIVER_CITY] = sReceiverCity;
                    dr[btw.RECEIVER_COMPANY_NAME2] = sReceiverCompanyName;
                    dr[btw.RECEIVER_COUNTRY] = sReceiverCountry;
                    dr[btw.RECEIVER_EMAIL] = sReceiverContactEmail;
                    dr[btw.RECEIVER_FAX] = sReceiverContactFax;
                    dr[btw.RECEIVER_MOBILE] = sReceiverContactMobile;
                    dr[btw.RECEIVER_NAME] = sReceiverContactName;
                    dr[btw.RECEIVER_PHONE] = sReceiverContactPhone;
                    dr[btw.RECEIVER_ZIP_CODE] = sReceiverZIP_CODE;
                    dr[btw.RELEASE_NO] = _RELEASE_NO;
                    dr[btw.ROUND_NET_WEIGHT] = sROUND_NET_WEIGHT;
                    dr[btw.SAFETY_CUE_J] = string.Empty;
                    dr[btw.SAFETY_STREAM] = string.Empty;
                    dr[btw.SEQ_NUM] = sSEQ_NUM; // počet skládaček v obalu
                    dr[btw.SMALL_PACKAGE_QTY] = iSMALL_PACKAGE_QTY.ToString();
                    dr[btw.STATION_NUM] = sTESTING_STATION_NUMBER;
                    dr[btw.ST_LENGTH] = sST_LENGTH;
                    dr[btw.SUPPLY_WIRE_MATERIAL] = sSUPPLY_WIRE_MATERIAL;
                    dr[btw.TESTING_PROTOCOL_NUMBER] = sTESTING_PROTOCOL_NUMBER;
                    dr[btw.TESTING_STATION_NUMBER] = sTESTING_STATION_NUMBER;
                    dr[btw.TITLE] = sTitle;
                    dr[btw.TOTAL_PACKAGE_QTY] = sPACKAGE_QTY;
                    dr[btw.TOTAL_RESIST] = string.Empty;
                    dr[btw.UNIQUE_PACKAGE_ID] = string.Empty;
                    dr[btw.UNIQUE_SUBPACKAGE_ID] = string.Empty;
                    dr[btw.UN_NO] = sUN;
                    dr[btw.VENDOR_ADDRESS_ID] = string.Empty;
                    dr[btw.VENDOR_CITY] = string.Empty;
                    dr[btw.VENDOR_COMPANY] = string.Empty;
                    dr[btw.VENDOR_COUNTRY] = string.Empty;
                    dr[btw.VENDOR_NAME] = string.Empty;
                    dr[btw.VENDOR_STREET] = string.Empty;
                    dr[btw.VENDOR_ZIP_CODE] = string.Empty;
                    dr[btw.WIRE_ISULATOR_MATERIAL] = string.Empty;
                    dr[btw.WIRE_LENGTH] = string.Empty;
                    dr[btw.WIRE_LENGTH_TRACE] = string.Empty;
                    dr[btw.WORKER1] = string.Empty;
                    dr[btw.WORKER2] = string.Empty;


                    _dtNalepky.Rows.Add(dr);
                    for (int jj = 1; jj <= iSEQ_NUM; jj++)
                    {
                    }
                }
                //
                //string sTime = String.Format("{0}-{1}-{2} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                string sSQLp = "INSERT INTO nela(ZO_BatchNo,VZ_BatchNo,Cesta,DatumExportu,Uzivatel,Nela) VALUES(" +
                            "'" + _ZO_BatchNo + "'," +
                            "'" + _VZ_BatchNo + "'," +
                            "'" + sFile.Trim() + "'," +
                            "'" + sTime + "','" + _UserName.Trim() + "'," +
                            "'" + _NELA.Trim() + "'" +
                            ")";
                //(String.Format("yyyy-MM-dd HH:mm:ss",DateTime.Now) + ("\'" + ")")))))))))))));
                if (MSSQL_Execute(sSQLp) == false)
                {
                    MessageBox.Show("fce: MSSQL_Execute neproběhla v pořádku.");
                }
                Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
                cn.Close();
                sw.Dispose();
                sw.Close();
                MessageBox.Show("Soubor je vytvořen.");
                KopirujNalepky(true);
            }
            catch (Exception ex)
            {
                if (sw != null)
                {
                    sw.Close();
                }
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }
            this.Cursor = Cursors.Default;
        }

        private void buttonFtpCopyNaEstar_Click(object sender, EventArgs e)
        {

        }

        private void KopirujTXT_NA_IMT_FTP()
        {
            //
            //
            this.Cursor = Cursors.WaitCursor;
            FtpClient ftp = new FtpClient();
            if (ftp.ftp_file_exist(Globals.IMT_FTP_SERVER_IMT, "imt", "imt123456", "MEMCARD1/RecipesADR", Globals.IMT_FILE_RECEPTURA) == true)
            {
                MessageBox.Show("Soubor už je ve schválených");
                this.Cursor = Cursors.Default;
                return;
            }
            if (ftp.ftp_copy(Globals.IMT_FTP_SERVER, "austin", "austin123456", "zakazky", Globals.IMT_FOLDERFILE_RECEPTURA) == false)
            {
                MessageBox.Show("Kopírování na FTP se nezdařilo");
                this.Cursor = Cursors.Default;
                return;
            }
            // zapis do logu
            string sTime = String.Format("{2}.{1}.{0} {3}:{4}:{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

            string sSQLp = "Update nela SET DatumNaImtFtp = " +
                        "'" + sTime + "'" +
                        " WHERE ID =" + Globals.ID_TABLE_NELA;

            //(String.Format("yyyy-MM-dd HH:mm:ss",DateTime.Now) + ("\'" + ")")))))))))))));
            if (MSSQL_Execute(sSQLp) == false)
            {
                MessageBox.Show("fce: MSSQL_Execute neproběhla v pořádku.");
            }
            Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);            //
            MessageBox.Show("Soubor: " + Environment.NewLine + Globals.IMT_FILE_RECEPTURA + Environment.NewLine + 
                "byl zkopírován na IMT - FTP.","FTP - COPY");
            this.Cursor = Cursors.Default;
        }
        private void toolStripMenuItem_IMT_FTP_Click(object sender, EventArgs e)
        {
            if (dataGridViewSoubory.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vyberte řádek");
                return;
            }

            DataGridViewRow currentRow = dataGridViewSoubory.SelectedRows[0];

            string sMachine = Convert.ToString(currentRow.Cells["Nela"].Value);
            Globals.ID_TABLE_NELA = Convert.ToInt32(currentRow.Cells["ID"].Value);

            if (sMachine != NelaFile_MachineName.E3PR)
            {
                MessageBox.Show("Není vybrán ESTAR !!");
                return;
            }



            string sFileName = Convert.ToString(currentRow.Cells["Cesta"].Value);
            string sZO_BatchNo = Convert.ToString(currentRow.Cells["ZO_BatchNo"].Value);

            if (!File.Exists(sFileName))
            {
                MessageBox.Show("Soubor neexistuje" + Environment.NewLine + sFileName, "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            Globals.IMT_FILE_RECEPTURA = Path.GetFileName(sFileName);
            //
            string ext = Path.GetExtension(Globals.IMT_FILE_RECEPTURA);
            string sPom = sZO_BatchNo; //Globals.IMT_FILE_RECEPTURA.Replace(ext, "");
            string[] words = sPom.Split('-');
            _Customer_order = words[0];
            _LINE_NO = words[1];
            _RELEASE_NO = words[2];
            //
            Globals.IMT_FOLDERFILE_RECEPTURA = sFileName;
            //
            //KopirujTXT_NA_IMT_FTP();
            KopirujTXTaNalepky_NA_IMT_FTP(true);
            ;

        }
        private bool  NaplnListBtw()
        {
            this.Cursor = Cursors.WaitCursor;

            string sSQLp = string.Empty;

            bool bOk = false;
            string sFile = string.Empty;
            _Labels_BTW_names.Clear();

            string sCnString = "Data Source=" + Globals.TNSNAME+ ";UID=LOTYL;PWD=123456";

            if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
            {
                sSQLp = Globals.SQL_Nalepky_From_DMS_IFS9 + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'";
            }
            else
            {
                sSQLp = Globals.SQL_Nalepky_From_PO + _Customer_order + "^" + _LINE_NO + "^" + _RELEASE_NO + "^%'";
            }

            try
            {
                System.Data.OracleClient.OracleConnection conn = new System.Data.OracleClient.OracleConnection(sCnString);

                conn.Open();
                System.Data.OracleClient.OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQLp;
                System.Data.OracleClient.OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            switch (reader[0].GetType().ToString())
                            {
                                case "System.Decimal":
                                    sFile = Convert.ToString(reader.GetDecimal(0)); break;
                                case "System.DateTime":
                                    sFile = Convert.ToString(reader.GetDateTime(0).ToShortDateString()); break;
                                default:
                                    sFile = reader.GetString(0).ToString(); break;
                            }
                            if (sFile.ToUpper().Contains("-KTN") || sFile.ToUpper().Contains("-STN") || sFile.ToUpper().Contains("-PRO"))
                            {
                                if (sFile.ToUpper().Contains(".JPG"))
                                { }
                                else
                                { _Labels_BTW_names.Add(sFile); }
                                
                            }

                        }
                    }
                    bOk = true;

                }
                else
                {
                    bOk = false;
                    MessageBox.Show("Nálepky nejsou nastaveny v IFS !!");
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                bOk = false;
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
            return bOk;
        }

        private void dataGridViewSoubory_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewSoubory.SelectedRows.Count > 0)
            {
                DataGridViewRow currentRow = dataGridViewSoubory.SelectedRows[0];
                _ZO_BatchNo = Convert.ToString(currentRow.Cells[4].Value);
                _DOC_REFERENCE_OBJECT_KEY_VALUE = _ZO_BatchNo.Replace("-", "^");
            }

        }

        private void pictureBox_XML_Info_Click(object sender, EventArgs e)
        {
            frmInfo f = new frmInfo();
            f.Text1 = NelaFile.Properties.Settings.Default.Info;
            f.OpenFile();
            f.ShowDialog();
            f.Dispose();

        }

        private void checkBoxDecimalDelititerChange_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDecimalDelititerChange.Checked)
            {
                Globals.DECIMAL_SEPARATOR = DecimalSeparator.Dot;

            }
            else
            {
                Globals.DECIMAL_SEPARATOR = DecimalSeparator.Comma;
            }
        }

        private void buttonNajdiANaplnZO_Click(object sender, EventArgs e)
        {
            string sZO = textBoxFindVZ.Text.ToUpper();
            if (sZO == string.Empty)
            {
                MessageBox.Show("Není zadána Prodejní objednávka. V... !");
                return;
            }
            Napln_dataGridViewVZ(sZO);

        }

        private void checkBoxIFSVersion_CheckedChanged(object sender, EventArgs e)
        {
            dataGridViewVZ.DataSource = null;
            dataGridViewNalepky.DataSource = null;
            if (checkBoxIFSVersion.Checked)
            {
                Globals.TNSNAME = TNS_Names.AD9UPG3;
                Globals.IFS_VERSION = IFS_Versions.IFS_9;
                tsslDB.ForeColor = Color.Red;
            }
            else
            { 
                Globals.TNSNAME = TNS_Names.AD9PROD;
                Globals.IFS_VERSION = IFS_Versions.IFS_9;
                tsslDB.ForeColor = Color.Black;
            }
            textBoxNalepka.Text = string.Empty;
            tsslDB.Text = Globals.TNSNAME;
        }

        private void toolStripMenuItemNapln_Click(object sender, EventArgs e)
        {
            Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
        }
        private bool GrantAccess(string pFullPath, string pPath)
        {
            try
            { 
                DirectoryInfo dInfo = new DirectoryInfo(pPath);
                DirectorySecurity dSecurity = dInfo.GetAccessControl();
                dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), 
                    FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, 
                    AccessControlType.Allow));
                dInfo.SetAccessControl(dSecurity);

                if (File.Exists(pFullPath))
                { 
                    DirectoryInfo dirInfo = new DirectoryInfo(pFullPath);
                    DirectorySecurity dirSec = dirInfo.GetAccessControl();
                    dirSec.AddAccessRule(new FileSystemAccessRule("AUSTIN\\" + _UserName, FileSystemRights.FullControl, AccessCo‌ntrolType.Allow));
                    //dirSec.AddAccessRule(new FileSystemAccessRule("Users", FileSystemRights.FullControl, AccessCo‌ntrolType.Allow));
                    dirInfo.SetAccessControl(dirSec);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "GrantAccess", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void checkBoxFiltrVlastniSoubory_CheckedChanged(object sender, EventArgs e)
        {
            Napln_dataGridViewSoubory(checkBoxFiltrVlastniSoubory.Checked);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridViewVZ.DataSource = null;
            dataGridViewNalepky.DataSource = null;
            Globals.IFS_VERSION = IFS_Versions.IFS_9;
            if (radioButton1.Checked)
            {
                Globals.TNSNAME = TNS_Names.AD9PROD;
                tsslDB.ForeColor = Color.Black;
            }
            if (radioButton2.Checked)
            {
                Globals.TNSNAME = TNS_Names.AD9TEST;
                tsslDB.ForeColor = Color.Red;
            }
            if (radioButton3.Checked)
            {
                Globals.TNSNAME = TNS_Names.AD9UPG3;
                tsslDB.ForeColor = Color.Red;
            }
            textBoxNalepka.Text = string.Empty;
            tsslDB.Text = Globals.TNSNAME;
        }

        private void ToolStripMenuItem_Main_Elektra_Click(object sender, EventArgs e)
        {
            frmElektraData ff = new frmElektraData();

            ff.Order_No = _VZ_Order_No;
            ff.Release_No = _VZ_Release_No;
            ff.Sequence_No = _VZ_sequence_No;
            ff.Show();

        }

        private void buttonHledejNalepky_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewVZ.Rows)
            {
                string sOrder_no = row.Cells["Order_no"].Value.ToString();

            }
        }
        private string XML_ExLabeling_Get_Part_no(string pOrder_no, string pPart_no)
        {
            // Funkce f = new Funkce();
            string sPart_no = string.Empty;

            //string sSQL_Doc_no = "SELECT doc_no FROM doc_title_access WHERE doc_class = 'ZO-XML' AND doc_no like '%" + pOrder_no + "%'";

            string sSQL_Get_XML_BLOB =
               "SELECT file_data from IFSAPP.EDM_FILE_STORAGE a where a.doc_class = 'ZO-XML' and a.doc_no LIKE '%" + pOrder_no + "%'";


            try
            {
                using (var conn = new Oracle.ManagedDataAccess.Client.OracleConnection(Globals.ConString))
                {
                    conn.Open();
                
                
                    using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(sSQL_Get_XML_BLOB, conn))
                    {
                        //cmd.Parameters.Add(new OracleParameter("id", 1));
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (reader["file_data"] == DBNull.Value) return string.Empty;

                                byte[] blobBytes = (byte[])reader["file_data"];
                                var xmlDoc = new XmlDocument();
                                using (var ms = new MemoryStream(blobBytes))
                                {
                                    ms.Position = 0;
                                    xmlDoc.Load(ms); // correct overload for Stream
                                }
                                sPart_no = XML_GetElement_SID(xmlDoc, pPart_no);

                                // hotovo 👍
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return sPart_no;
        }
        public string XML_GetElement_SID(XmlDocument pXML, string pPart_no)
        {
            try
            {
                // Test
                XPathNavigator nav = pXML.CreateNavigator();
                XPathNodeIterator Iterator = nav.Select("Shipment/SummaryItems/SummaryItem");
                foreach (XPathNavigator book in Iterator)
                {
                    XPathNavigator nv = book.SelectSingleNode("ProducerProductCode");
                    string sBN = nv == null ? string.Empty : nv.Value;
                    if (sBN.Contains(pPart_no) || pPart_no.Contains(sBN))
                    {
                        return sBN;
                    }

                    //XPathNavigator nvSID = book.SelectSingleNode("@SID");
                    //string s_SID = nvSID == null ? string.Empty : nvSID.Value;
                    //if (s_SID == pSID)
                    //{
                    //    return sBN;
                    //}

                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }

        }

    }
}
