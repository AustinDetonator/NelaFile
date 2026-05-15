using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Diagnostics;
using BarTender;
using Seagull.BarTender.Print;

namespace NelaFile
{
    public partial class frmNalepkyTisk : Form
    {
        #region Public
        public string Document_code { get; set; }
        public string PackageBarcodeID { get; set; }
        public string SPZ { get; set; }
        public string ConnString { get; set; }
        public string ConString_MSSQL_Austin_Trc { get; set; }

        public string UserIFS { get; set; }
        public string ComputerName { get; set; }
        public string Version { get; set; }
        public string Filtr { get; set; }
        public DataSet dsNalepky;
        //
        public string ImageFileName = string.Empty;
        public List<string> ColumnListLabels = new List<string>();

        //
        #endregion Public
        #region MistniPromenne
        private Bitmap _Bitmap;
#pragma warning disable CS0169 // The field 'frmNalepkyTisk._Components' is never used
        private IContainer _Components;
#pragma warning restore CS0169 // The field 'frmNalepkyTisk._Components' is never used
#pragma warning disable CS0169 // The field 'frmNalepkyTisk._PictureBoxMain' is never used
        private PictureBox _PictureBoxMain;
#pragma warning restore CS0169 // The field 'frmNalepkyTisk._PictureBoxMain' is never used
        //
        private int _Dgw_Current_Row = -1;
        private string _BTW_TYP = string.Empty;
        private string _FileNameBTW = string.Empty;
        private string _FileNameFolderBTW = string.Empty;
        private string _DOC_CLASS = string.Empty;

        private DataTable _dtNalepky;
        DataColumn[] _keysNalepky = new DataColumn[5];
        #endregion

        public frmNalepkyTisk()
        {
            InitializeComponent();
            labelInfoZO.Text = ZO.ORDER_NO.Trim() + "-" + ZO.LINE_NO.Trim() + "-" + ZO.RELEASE_NO.Trim();
            labelCustomer.Text = ZO.Customer.Trim();
            //textBoxInfoZO.Text = ZO.ORDER_NO + " " + ZO.LINE_NO + " " + ZO.RELEASE_NO;
            PrintersGet();
        }
        private void dataGridViewBTW_Napln()
        {
            this.Cursor = Cursors.WaitCursor;

            dataGridViewBTW.RowHeadersWidth = 20;
            dataGridViewBTW.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewBTW.AllowUserToResizeRows = false;
            dataGridViewBTW.MultiSelect = true;
            dataGridViewBTW.AllowUserToResizeColumns = false;
            dataGridViewBTW.AllowUserToAddRows = false;

            DataTable dt = this.dsNalepky.Tables[0];
            //dgZakazky.DisplayMember = dt.Columns("doc_no").ToString
            dataGridViewBTW.BindingContext = new BindingContext();
            dataGridViewBTW.DataSource = dt;
            //
            dataGridViewBTW.Columns["FILE_NAME"].HeaderText = "FILE_NAME";
            dataGridViewBTW.Columns["FILE_NAME"].Width = 250;
            dataGridViewBTW.Columns["DOC_CLASS"].HeaderText = "DOC_CLASS";
            dataGridViewBTW.Columns["DOC_CLASS"].Visible = false;

            //CheckedListBox1.DisplayMember = "doc_no"
            if (dataGridViewBTW.CurrentRow != null)
            {
                //textBoxNalepka.Text = "\\\\Bap01\\IFS_NALEPKY\\" + this.dataGridViewNalepky[this.dataGridViewNalepky.CurrentRow.Index, 0].Value.ToString();
                //textBoxNalepka.Text = _FOLDER_NALEPKY + this.dataGridViewNalepky[this.dataGridViewNalepky.CurrentRow.Index, 0].Value.ToString();
            }
            this.Cursor = Cursors.Default;
        }

        private void frmNalepkyTisk_Load(object sender, EventArgs e)
        {
            dataGridViewBTW_Napln();
        }

        private void btnLabelPreview_Click(object sender, EventArgs e)
        {
            if (dataGridViewLabels.Rows.Count > 0)
            {
                StreamWriter swLog_NelaFile = new StreamWriter(Globals.NELAFILE_LogFile, true, System.Text.Encoding.Default);
                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " btnLabelPreview_Click()");
                swLog_NelaFile.Close();
                NalepkaNahled(false, dataGridViewLabels.CurrentRow.Index);
            }
            else
            {
                MessageBox.Show("Nejsou načteny řádky pro nálepky !");
            }
        }
        private void NalepkaNahled(bool bPrint,int pRow )
        {
            //ApplicationClass BTApp = (ApplicationClass)null;
            StreamWriter swLog_NelaFile = new StreamWriter(Globals.NELAFILE_LogFile, true, System.Text.Encoding.Default);
            swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() " + bPrint.ToString());
            
            //
            //pictureBoxLabelPreview.Image = null;
            //try
            //{
            //    pictureBoxLabelPreview.Image.Dispose();
            //} 
            //catch (Exception e)
            //{
            //    swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() " + e.Message.ToString());
            //}

            swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() " + _FileNameFolderBTW);
            if (!File.Exists(_FileNameFolderBTW))
            {
                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() " + Environment.NewLine + " Soubor : " + _FileNameFolderBTW + " neexistuje !!");
                swLog_NelaFile.Close();
                MessageBox.Show("Soubor :" + Environment.NewLine + _FileNameFolderBTW + Environment.NewLine + " neexistuje !!");
                return;
            }
            if (Path.GetExtension(_FileNameFolderBTW).ToLower() == ".jpg")
            {
                //MessageBox.Show("Soubor :" + Environment.NewLine + _FileNameFolderBTW + Environment.NewLine + " není správného typu BTW !!");
                //this._Bitmap = new Bitmap(_FileNameFolderBTW);
                //Bitmap btmpJPG = new Bitmap(_FileNameFolderBTW);
                //pictureBoxLabelPreview.Image = (Image)this._Bitmap;
                swLog_NelaFile.Close();

                frmPreviewLabel f = new frmPreviewLabel();
                f.PictureFolderFile = _FileNameFolderBTW;
                f.ShowDialog();
                f.Dispose();
                return;
            }
            if (Path.GetExtension(_FileNameFolderBTW).ToLower() == ".pdf")
            {
                //MessageBox.Show("Soubor :" + Environment.NewLine + _FileNameFolderBTW + Environment.NewLine + " není správného typu BTW !!");
                System.Diagnostics.Process.Start(_FileNameFolderBTW);

                //this._Bitmap = new Bitmap(_FileNameFolderBTW);
                //pictureBoxLabelPreview.Image = (Image)this._Bitmap;
                swLog_NelaFile.Close();
                return;
            }
            if (Path.GetExtension(_FileNameFolderBTW).ToLower() == ".zpl")
            {
                MessageBox.Show("Soubor typu ZPL nejde zobrazit.");
                swLog_NelaFile.Close();
                return;
            }

            //string str1 = System.Windows.Forms.Application.StartupPath + (object)'\\';

            BarTender.Format format = (BarTender.Format)null;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                DateTime now = DateTime.Now;
                //
                // Declare a BarTender application variable
                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled()");

                BarTender.Application btApp;
                BarTender.Format btFormat;
                btApp = new BarTender.Application();
                //public static ApplicationClass BTApp = (ApplicationClass)null;
                btApp.Visible = false;

                // Open a BarTender document
                btFormat = btApp.Formats.Open(_FileNameFolderBTW, true, "");
                btFormat.Printer = Globals.PRINTER;
                // netiskne to na A4
                BarTender.PageSetup btPageSetup;
                btPageSetup = btFormat.PageSetup;
                btPageSetup.Orientation = BarTender.BtOrientation.btPortrait;
                //
                //btApp.Quit(BarTender.BtSaveOptions.btDoNotSaveChanges);
                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled()" + btFormat.FileName.ToString());

                // Export document to a print preview

                BarTender.BtPrintResult res;
                //string sPictureFile = @"\nahled%PageNumber%.bmp";
                string sPictureFile = @"\nahled" + DateTime.Now.Millisecond + ".jpg";
                string sPictureFolderFile = Globals.FOLDER_BTW_EXPORT + sPictureFile;
                Globals.JPG_LAST_FOLDER_FILENAME = sPictureFolderFile;

                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled()" + sPictureFolderFile.ToString());

                //"Label_%PageNumber%_Preview.jpg"
                //S:\KOPY_archiv_vykresova dokumentace\VYKRESY_B A L E N I\aBARTENDER_v901_Nova grafika_data\austin_sablona_1.txt
                //btFormat.Databases.GetDatabase("tisk_data.txt");
                //BtwRow dataBoundItem = (BtwRow)dataGridViewLabels.CurrentRow.DataBoundItem;
                //
                string sHeader = btFormat.get_UsedNames(BtNameType.btUsedDatabaseFields, ";");
                if (sHeader != string.Empty)
                {
                    //"austin_sablona_1.austin_sablona_1.13_PACKAGE_NUM2;austin_sablona_1.austin_sablona_1.14_CATALOG_NO;austin_sablona_1.austin_sablona_1.36_PACKAGE_QTY;austin_sablona_1.austin_sablona_1.37_SMALL_PACKAGE_QTY;austin_sablona_1.austin_sablona_1.38_PACKAGE_GROSS_WEIGHT;austin_sablona_1.austin_sablona_1.40_PACKAGE_NET_WEIGHT;austin_sablona_1.austin_sablona_1.41_EXPLOSIVE_NET_WEIGHT;austin_sablona_1.austin_sablona_1.42_UN_NO;austin_sablona_1.austin_sablona_1.43_ADR_RID_CLASS;austin_sablona_1.austin_sablona_1.52_SUPPLY_WIRE_MATERIAL;austin_sablona_1.austin_sablona_1.53_WIRE_ISULATOR_MATERIAL;austin_sablona_1.austin_sablona_1.54_WIRE_LENGTH;austin_sablona_1.austin_sablona_1.57_NOMINAL_TIME_EA;austin_sablona_1.austin_sablona_1.61_BATCH_NO;austin_sablona_1.austin_sablona_1.62_TESTING_STATION_NUMBER;austin_sablona_1.austin_sablona_1.63_TESTING_PROTOCOL_NUMBER;austin_sablona_1.austin_sablona_1.69_DATE;austin_sablona_1.austin_sablona_1.70_EXPIRATION;austin_sablona_1.austin_sablona_1.85_PRODUCER"
                    string sTable = sHeader.Substring(0, sHeader.IndexOf("."));
                    string sFileData = Globals.FOLDER_BTW_EXPORT + @"\tisk_data.txt";
                    StreamWriter sw = new StreamWriter(sFileData, false, System.Text.Encoding.Default);
                    string sLine1 = string.Empty;
                    string sLine2 = string.Empty;
                    string sPACKAGE_BARCODE_ID = dataGridViewLabels[BtwHdr.c34_UNIQUE_PACKAGE_ID, dataGridViewLabels.CurrentRow.Index].Value.ToString();
                    string sSUBPACKAGE_BARCODE_ID = dataGridViewLabels[BtwHdr.c35_UNIQUE_SUBPACKAGE_ID, dataGridViewLabels.CurrentRow.Index].Value.ToString();
                    string s96_DMX_PRODUCER_Prefix = "=90"; //«GS»

                    for (int i = 2; i < dataGridViewLabels.ColumnCount; i++) // první 2 sloupce mám pomocné - číslo kartonu a skládačky, aby byly na prvním místě
                    {
                        string sSeparator = ((i == dataGridViewLabels.ColumnCount - 1) ? string.Empty : ";");
                        string sHeaderColumn = dataGridViewLabels.Columns[i].HeaderText;
                        sLine1 = sLine1 + sHeaderColumn + sSeparator;
                        // index řádku : dataGridViewLabels.CurrentRow.Index
                        string sValue = dataGridViewLabels[i, pRow].Value.ToString();
                        if (dataGridViewLabels.Columns[i].HeaderText == "96_DMX_PRODUCER")                     //96_DMX_PRODUCER
                        {
                            sValue = "90" + sValue + (char)29;
                        }
                        sLine2 = sLine2 + sValue + sSeparator;
                        // get ColumnName from BTW  (from Header)
                        string sPrefix = strSplit(sHeader, sHeaderColumn, ";");
                        //string sColumnNameBTW = "Textový soubor 1.austin_sablona_1." + sHeaderColumn;
                        string sColumnNameBTW = sPrefix +  sHeaderColumn;
                        if (sPrefix != string.Empty)
                        {
                            btFormat.set_UsedNamesSampleData(BarTender.BtUsedNamesSampleDataType.btUsedDatabaseFieldsSampleData, sColumnNameBTW, sValue);
                        }
                        //btFormat.SetNamedSubStringValue("KsVKartonu","100");
                    }
                    sw.WriteLine(sLine1 + ";");
                    sw.WriteLine(sLine2 + ";");
                    sw.Close();
                    string index = sTable;
                    object obj = index;
                    btFormat.Databases.GetDatabase(obj).TextFile.FileName = sFileData;

                }

                // new with engine
                //Engine btwEngine = new Engine();
                //btwEngine.Start();
                //LabelFormatDocument btwFormat = btwEngine.Documents.Open(_FileNameFolderBTW);
                //btwFormat.DatabaseConnections["Textový soubor 1"].FileName = sFileData;
                //Seagull.BarTender.Print.Messages bteMessages;
                //btwFormat.ExportPrintPreviewToFile(Globals.FOLDER_BTW_EXPORT, sPictureFile, ImageType.BMP,Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300),Color.AliceBlue, OverwriteOptions.DoNotOverwrite, true, true, out bteMessages);
                ////btwFormat.Print();
                //string messageString = "\n\nMessages:";

                //foreach (BarTender.Message message in bteMessages)
                //{
                //    messageString += "\n\n" + message.Message.ToString();
                //}

                //btwEngine.Stop();
                //
                BarTender.Messages btMessages;
                //string index = "Textový soubor 1";

                // ***** print
                if (!bPrint)
                {
                    btFormat.PrintToFile = true;
                }
                else
                {
                }
                //BarTender.PrinterCodeTemplate bt = btFormat.PrinterCodeTemplate;
                btFormat.Print("Nálepka", false, -1, out btMessages);
                // *****

                //btFormat.Databases.GetDatabase((object)1).TextFile.FileName = sFileData;
                //austin_sablona_1.txt
                //string index = "Textový soubor 1";
                //object obj = index;
                //BarTender.Messages btMessages;
                //res = btFormat.ExportPrintPreviewToImage(Globals.FOLDER_BTW_EXPORT, sPictureFile, "bmp", BarTender.BtColors.btColors24Bit, 200, 13117215, BarTender.BtSaveOptions.btSaveChanges, true, true, out btMessages);
                ////
                //messageString = "\n\nMessages:";
                //foreach (BarTender.Message message in btMessages)
                //{
                //    messageString += "\n\n" + message.Message.ToString();
                //}

                btFormat.ExportToFile(sPictureFolderFile,"jpg", BtColors.btColors256, BtResolution.btResolutionScreen, BtSaveOptions.btDoNotSaveChanges);

                swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() - after btFormat.ExportToFile " + sPictureFolderFile.ToString());

                // ***** print
                if (!bPrint)
                {
                    frmPreviewLabel f = new frmPreviewLabel();
                    f.PictureFolderFile = sPictureFolderFile;
                    f.ShowDialog();
                    f.Dispose();
                }

                //Bitmap btmp = new Bitmap(sPictureFolderFile);

                //swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() - after this._Bitmap = new Bitmap(sPictureFolderFile) " + sPictureFolderFile);

                //pictureBoxLabelPreview.Image = (Image)btmp; //this._Bitmap;

                //swLog_NelaFile.WriteLine(DateTime.Now.ToString() + " NalepkaNahled() - after pictureBoxLabelPreview.Image = (Image)this._Bitmap " + sPictureFolderFile);

                btFormat.Close(BtSaveOptions.btDoNotSaveChanges);
                btApp.Quit(BarTender.BtSaveOptions.btDoNotSaveChanges);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(btApp);
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                int num2 = (int)MessageBox.Show(ex.ToString());
                MessageBox.Show(ex.Message.ToString());
                //Globals.LogError(nameof(btnPreview_Click), "Chyba: " + ex.ToString());
            }
            finally
            {
                if (format != null)
                    format.Close(BtSaveOptions.btDoNotSaveChanges);
                this.Cursor = Cursors.Default;
            }
            swLog_NelaFile.Close();
            //
            //////this.Width = this._Bitmap.Width + 5;
            //////this.Height = this._Bitmap.Height + 40;
            ////pictureBoxLabelPreview.Image = (Image)this._Bitmap;
        }

        private void dataGridViewBTW_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewBTW.CurrentRow == null)
            {
                return;
            }
            //pictureBoxLabelPreview.Image = null;pictureBoxLabelPreview.Refresh();
            _Dgw_Current_Row = dataGridViewBTW.CurrentRow.Index;

            _DOC_CLASS = dataGridViewBTW[1, _Dgw_Current_Row].Value.ToString();
            _FileNameBTW = dataGridViewBTW[0, _Dgw_Current_Row].Value.ToString();
            _FileNameFolderBTW = @Globals.FOLDER_BTW_EXPORT + @"\" + _FileNameBTW;



            string[] keys = new string[] {"KTN","STN","KOS","KUN","BTU","SBS"};

            string sKeyResult = keys.FirstOrDefault<string>(s => _FileNameBTW.Contains(s));

            switch (sKeyResult)
            {
                case BtwType.KTN:
                    _BTW_TYP = BtwType.KTN;
                    break;
                case BtwType.STN:
                    _BTW_TYP = BtwType.STN;
                    break;
                default:
                    _BTW_TYP = BtwType.SIGNO;
                    break;
            }
            //Napln_dataGridViewNalepky();
            if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
            {
                //NalepkyGeneruj_IFS9();
            }
            else
            {
                //NalepkyGeneruj();
            }
        }
        private void Napln_dataGridViewLabels()
        {
            //  uz udelane XML pro pY.listy
#pragma warning disable CS0219 // The variable 'cnsSQL01' is assigned but its value is never used
            string cnsSQL01 = "Persist Security Info=True;User ID=epo;Password=123456;Initial Catalog=epo;Data Source=SQL01";
#pragma warning restore CS0219 // The variable 'cnsSQL01' is assigned but its value is never used
            //SqlConnection cnSQL01;
            //SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            //cnSQL01 = new SqlConnection(cnsSQL01);
            //cnSQL01.Open();
#pragma warning disable CS0219 // The variable 'sSQLp' is assigned but its value is never used
            string sSQLp = "SELECT DatumExportu,Nela,Uzivatel,ZO_BatchNo,VZ_BatchNo,Cesta FROM nela ORDER BY DatumExportu desc";
#pragma warning restore CS0219 // The variable 'sSQLp' is assigned but its value is never used
            //SqlCommand cmd = new SqlCommand(sSQLp, cnSQL01);
            // 
            //adapter.SelectCommand = cmd;
            //adapter.Fill(ds, "Nela");
            DataTable dt = ds.Tables[0];
            // dgZakazky.DisplayMember = dt.Columns("doc_no").ToString
            // dgZakazky.BindingContext = New BindingContext
            dataGridViewLabels.DataSource = dt;
            // CheckedListBox1.DisplayMember = "doc_no"



            //cnSQL01.Close();

            dataGridViewLabels.Columns["ID"].HeaderText = "Datum exportu";
            dataGridViewLabels.Columns["ID"].Width = 100;

            dataGridViewLabels.Columns[1].HeaderText = "NELA";
            dataGridViewLabels.Columns[1].Width = 70;

            dataGridViewLabels.Columns[2].HeaderText = "Uživatel";
            dataGridViewLabels.Columns[2].Width = 80;

            dataGridViewLabels.Columns[3].HeaderText = "ZO_BatchNo";
            dataGridViewLabels.Columns[3].Width = 120;

            dataGridViewLabels.Columns[4].HeaderText = "VZ_BatchNo";
            dataGridViewLabels.Columns[4].Width = 120;

            dataGridViewLabels.Columns[5].HeaderText = "Cesta+soubor";
            dataGridViewLabels.Columns[5].Width = 400;





        }

        private void buttonCreateDataForLabels_Click(object sender, EventArgs e)
        {

        }

        private void buttonLabelsCreate_Click(object sender, EventArgs e)
        {
            if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
            {
                NalepkyGeneruj_IFS9();
            }
            else
            {
                NalepkyGeneruj();
            }
        }
        private void MakeDataTable_Nalepky()
        {
            //DataTable newProducts = new DataTable("NewProducts");
            _dtNalepky = new DataTable();
            _keysNalepky = new DataColumn[1];

            // Add three column objects to the table. 
            //keys[0] = Station;
            // pomocný pro rozlišení typu nálepky

            //DataColumn BTW_TYPE = new DataColumn(); BTW_TYPE.DataType = System.Type.GetType("System.String"); BTW_TYPE.ColumnName = btw.BTW_TYPE; _dtNalepky.Columns.Add(BTW_TYPE);

            DataColumn PACKAGE_ID = new DataColumn(); PACKAGE_ID.DataType = System.Type.GetType("System.String"); PACKAGE_ID.ColumnName = "PACKAGE_ID"; _dtNalepky.Columns.Add(PACKAGE_ID);
            DataColumn SUBPACKAGE_ID = new DataColumn(); SUBPACKAGE_ID.DataType = System.Type.GetType("System.String"); SUBPACKAGE_ID.ColumnName = "SUBPACKAGE_ID"; _dtNalepky.Columns.Add(SUBPACKAGE_ID);


            DataColumn c01_DOC_CLASS = new DataColumn(); c01_DOC_CLASS.DataType = System.Type.GetType("System.String"); c01_DOC_CLASS.ColumnName = BtwHdr.c01_DOC_CLASS; _dtNalepky.Columns.Add(c01_DOC_CLASS);
            DataColumn c02_TITLE = new DataColumn(); c02_TITLE.DataType = System.Type.GetType("System.String"); c02_TITLE.ColumnName = BtwHdr.c02_TITLE; _dtNalepky.Columns.Add(c02_TITLE);
            DataColumn c03_DOC_NO = new DataColumn(); c03_DOC_NO.DataType = System.Type.GetType("System.String"); c03_DOC_NO.ColumnName = BtwHdr.c03_DOC_NO; _dtNalepky.Columns.Add(c03_DOC_NO);
            DataColumn c04_DOC_REV = new DataColumn(); c04_DOC_REV.DataType = System.Type.GetType("System.String"); c04_DOC_REV.ColumnName = BtwHdr.c04_DOC_REV; _dtNalepky.Columns.Add(c04_DOC_REV);
            DataColumn c05_FILE_NAME = new DataColumn(); c05_FILE_NAME.DataType = System.Type.GetType("System.String"); c05_FILE_NAME.ColumnName = BtwHdr.c05_FILE_NAME; _dtNalepky.Columns.Add(c05_FILE_NAME);
            DataColumn c06_CUSTOMER_PO_NO = new DataColumn(); c06_CUSTOMER_PO_NO.DataType = System.Type.GetType("System.String"); c06_CUSTOMER_PO_NO.ColumnName = BtwHdr.c06_CUSTOMER_PO_NO; _dtNalepky.Columns.Add(c06_CUSTOMER_PO_NO);
            DataColumn c07_ORDER_NO = new DataColumn(); c07_ORDER_NO.DataType = System.Type.GetType("System.String"); c07_ORDER_NO.ColumnName = BtwHdr.c07_ORDER_NO; _dtNalepky.Columns.Add(c07_ORDER_NO);
            DataColumn c08_LINE_NO = new DataColumn(); c08_LINE_NO.DataType = System.Type.GetType("System.String"); c08_LINE_NO.ColumnName = BtwHdr.c08_LINE_NO; _dtNalepky.Columns.Add(c08_LINE_NO);
            DataColumn c09_RELEASE_NO = new DataColumn(); c09_RELEASE_NO.DataType = System.Type.GetType("System.String"); c09_RELEASE_NO.ColumnName = BtwHdr.c09_RELEASE_NO; _dtNalepky.Columns.Add(c09_RELEASE_NO);
            DataColumn c10_LINE_ITEM_NO = new DataColumn(); c10_LINE_ITEM_NO.DataType = System.Type.GetType("System.String"); c10_LINE_ITEM_NO.ColumnName = BtwHdr.c10_LINE_ITEM_NO; _dtNalepky.Columns.Add(c10_LINE_ITEM_NO);
            DataColumn c11_PACKAGE_NUM1 = new DataColumn(); c11_PACKAGE_NUM1.DataType = System.Type.GetType("System.String"); c11_PACKAGE_NUM1.ColumnName = BtwHdr.c11_PACKAGE_NUM1; _dtNalepky.Columns.Add(c11_PACKAGE_NUM1);
            DataColumn c12_SEQ_NUM = new DataColumn(); c12_SEQ_NUM.DataType = System.Type.GetType("System.String"); c12_SEQ_NUM.ColumnName = BtwHdr.c12_SEQ_NUM; _dtNalepky.Columns.Add(c12_SEQ_NUM);
            DataColumn c13_PACKAGE_NUM2 = new DataColumn(); c13_PACKAGE_NUM2.DataType = System.Type.GetType("System.String"); c13_PACKAGE_NUM2.ColumnName = BtwHdr.c13_PACKAGE_NUM2; _dtNalepky.Columns.Add(c13_PACKAGE_NUM2);
            DataColumn c14_CATALOG_NO = new DataColumn(); c14_CATALOG_NO.DataType = System.Type.GetType("System.String"); c14_CATALOG_NO.ColumnName = BtwHdr.c14_CATALOG_NO; _dtNalepky.Columns.Add(c14_CATALOG_NO);
            DataColumn c15_CATALOG_DESC = new DataColumn(); c15_CATALOG_DESC.DataType = System.Type.GetType("System.String"); c15_CATALOG_DESC.ColumnName = BtwHdr.c15_CATALOG_DESC; _dtNalepky.Columns.Add(c15_CATALOG_DESC);
            DataColumn c16_VENDOR_COMPANY = new DataColumn(); c16_VENDOR_COMPANY.DataType = System.Type.GetType("System.String"); c16_VENDOR_COMPANY.ColumnName = BtwHdr.c16_VENDOR_COMPANY; _dtNalepky.Columns.Add(c16_VENDOR_COMPANY);
            DataColumn c17_VENDOR_NAME = new DataColumn(); c17_VENDOR_NAME.DataType = System.Type.GetType("System.String"); c17_VENDOR_NAME.ColumnName = BtwHdr.c17_VENDOR_NAME; _dtNalepky.Columns.Add(c17_VENDOR_NAME);
            DataColumn c18_VENDOR_ADDRESS_ID = new DataColumn(); c18_VENDOR_ADDRESS_ID.DataType = System.Type.GetType("System.String"); c18_VENDOR_ADDRESS_ID.ColumnName = BtwHdr.c18_VENDOR_ADDRESS_ID; _dtNalepky.Columns.Add(c18_VENDOR_ADDRESS_ID);
            DataColumn c19_VENDOR_STREET = new DataColumn(); c19_VENDOR_STREET.DataType = System.Type.GetType("System.String"); c19_VENDOR_STREET.ColumnName = BtwHdr.c19_VENDOR_STREET; _dtNalepky.Columns.Add(c19_VENDOR_STREET);
            DataColumn c20_VENDOR_CITY = new DataColumn(); c20_VENDOR_CITY.DataType = System.Type.GetType("System.String"); c20_VENDOR_CITY.ColumnName = BtwHdr.c20_VENDOR_CITY; _dtNalepky.Columns.Add(c20_VENDOR_CITY);
            DataColumn c21_VENDOR_ZIP_CODE = new DataColumn(); c21_VENDOR_ZIP_CODE.DataType = System.Type.GetType("System.String"); c21_VENDOR_ZIP_CODE.ColumnName = BtwHdr.c21_VENDOR_ZIP_CODE; _dtNalepky.Columns.Add(c21_VENDOR_ZIP_CODE);
            DataColumn c22_VENDOR_COUNTRY = new DataColumn(); c22_VENDOR_COUNTRY.DataType = System.Type.GetType("System.String"); c22_VENDOR_COUNTRY.ColumnName = BtwHdr.c22_VENDOR_COUNTRY; _dtNalepky.Columns.Add(c22_VENDOR_COUNTRY);
            DataColumn c23_RECEIVER_COMPANY_NAME2 = new DataColumn(); c23_RECEIVER_COMPANY_NAME2.DataType = System.Type.GetType("System.String"); c23_RECEIVER_COMPANY_NAME2.ColumnName = BtwHdr.c23_RECEIVER_COMPANY_NAME2; _dtNalepky.Columns.Add(c23_RECEIVER_COMPANY_NAME2);
            DataColumn c24_RECEIVER_ADDRESS1 = new DataColumn(); c24_RECEIVER_ADDRESS1.DataType = System.Type.GetType("System.String"); c24_RECEIVER_ADDRESS1.ColumnName = BtwHdr.c24_RECEIVER_ADDRESS1; _dtNalepky.Columns.Add(c24_RECEIVER_ADDRESS1);
            DataColumn c25_RECEIVER_ADDRESS2 = new DataColumn(); c25_RECEIVER_ADDRESS2.DataType = System.Type.GetType("System.String"); c25_RECEIVER_ADDRESS2.ColumnName = BtwHdr.c25_RECEIVER_ADDRESS2; _dtNalepky.Columns.Add(c25_RECEIVER_ADDRESS2);
            DataColumn c26_RECEIVER_ZIP_CODE = new DataColumn(); c26_RECEIVER_ZIP_CODE.DataType = System.Type.GetType("System.String"); c26_RECEIVER_ZIP_CODE.ColumnName = BtwHdr.c26_RECEIVER_ZIP_CODE; _dtNalepky.Columns.Add(c26_RECEIVER_ZIP_CODE);
            DataColumn c27_RECEIVER_CITY = new DataColumn(); c27_RECEIVER_CITY.DataType = System.Type.GetType("System.String"); c27_RECEIVER_CITY.ColumnName = BtwHdr.c27_RECEIVER_CITY; _dtNalepky.Columns.Add(c27_RECEIVER_CITY);
            DataColumn c28_RECEIVER_COUNTRY = new DataColumn(); c28_RECEIVER_COUNTRY.DataType = System.Type.GetType("System.String"); c28_RECEIVER_COUNTRY.ColumnName = BtwHdr.c28_RECEIVER_COUNTRY; _dtNalepky.Columns.Add(c28_RECEIVER_COUNTRY);
            DataColumn c29_RECEIVER_NAME = new DataColumn(); c29_RECEIVER_NAME.DataType = System.Type.GetType("System.String"); c29_RECEIVER_NAME.ColumnName = BtwHdr.c29_RECEIVER_NAME; _dtNalepky.Columns.Add(c29_RECEIVER_NAME);
            DataColumn c30_RECEIVER_PHONE = new DataColumn(); c30_RECEIVER_PHONE.DataType = System.Type.GetType("System.String"); c30_RECEIVER_PHONE.ColumnName = BtwHdr.c30_RECEIVER_PHONE; _dtNalepky.Columns.Add(c30_RECEIVER_PHONE);
            DataColumn c31_RECEIVER_FAX = new DataColumn(); c31_RECEIVER_FAX.DataType = System.Type.GetType("System.String"); c31_RECEIVER_FAX.ColumnName = BtwHdr.c31_RECEIVER_FAX; _dtNalepky.Columns.Add(c31_RECEIVER_FAX);
            DataColumn c32_RECEIVER_MOBILE = new DataColumn(); c32_RECEIVER_MOBILE.DataType = System.Type.GetType("System.String"); c32_RECEIVER_MOBILE.ColumnName = BtwHdr.c32_RECEIVER_MOBILE; _dtNalepky.Columns.Add(c32_RECEIVER_MOBILE);
            DataColumn c33_RECEIVER_EMAIL = new DataColumn(); c33_RECEIVER_EMAIL.DataType = System.Type.GetType("System.String"); c33_RECEIVER_EMAIL.ColumnName = BtwHdr.c33_RECEIVER_EMAIL; _dtNalepky.Columns.Add(c33_RECEIVER_EMAIL);
            DataColumn c34_UNIQUE_PACKAGE_ID = new DataColumn(); c34_UNIQUE_PACKAGE_ID.DataType = System.Type.GetType("System.String"); c34_UNIQUE_PACKAGE_ID.ColumnName = BtwHdr.c34_UNIQUE_PACKAGE_ID; _dtNalepky.Columns.Add(c34_UNIQUE_PACKAGE_ID);
            DataColumn c35_UNIQUE_SUBPACKAGE_ID = new DataColumn(); c35_UNIQUE_SUBPACKAGE_ID.DataType = System.Type.GetType("System.String"); c35_UNIQUE_SUBPACKAGE_ID.ColumnName = BtwHdr.c35_UNIQUE_SUBPACKAGE_ID; _dtNalepky.Columns.Add(c35_UNIQUE_SUBPACKAGE_ID);

            //DataColumn c34_UNIQUE_PACKAGE_ID = new DataColumn(); c34_UNIQUE_PACKAGE_ID.DataType = System.Type.GetType("System.String"); c34_UNIQUE_PACKAGE_ID.ColumnName = BtwHdr.c34_UNIQUE_PACKAGE_ID; _dtNalepky.Columns.Add(c34_UNIQUE_PACKAGE_ID);
            //DataColumn c35_UNIQUE_SUBPACKAGE_ID = new DataColumn(); c35_UNIQUE_SUBPACKAGE_ID.DataType = System.Type.GetType("System.String"); c35_UNIQUE_SUBPACKAGE_ID.ColumnName = BtwHdr.c35_UNIQUE_SUBPACKAGE_ID; _dtNalepky.Columns.Add(c35_UNIQUE_SUBPACKAGE_ID);
            DataColumn c36_PACKAGE_QTY = new DataColumn(); c36_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c36_PACKAGE_QTY.ColumnName = BtwHdr.c36_PACKAGE_QTY; _dtNalepky.Columns.Add(c36_PACKAGE_QTY);
            DataColumn c37_SMALL_PACKAGE_QTY = new DataColumn(); c37_SMALL_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c37_SMALL_PACKAGE_QTY.ColumnName = BtwHdr.c37_SMALL_PACKAGE_QTY; _dtNalepky.Columns.Add(c37_SMALL_PACKAGE_QTY);
            DataColumn c38_PACKAGE_GROSS_WEIGHT = new DataColumn(); c38_PACKAGE_GROSS_WEIGHT.DataType = System.Type.GetType("System.String"); c38_PACKAGE_GROSS_WEIGHT.ColumnName = BtwHdr.c38_PACKAGE_GROSS_WEIGHT; _dtNalepky.Columns.Add(c38_PACKAGE_GROSS_WEIGHT);
            DataColumn c39_ROUND_NET_WEIGHT = new DataColumn(); c39_ROUND_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c39_ROUND_NET_WEIGHT.ColumnName = BtwHdr.c39_ROUND_NET_WEIGHT; _dtNalepky.Columns.Add(c39_ROUND_NET_WEIGHT);
            DataColumn c40_PACKAGE_NET_WEIGHT = new DataColumn(); c40_PACKAGE_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c40_PACKAGE_NET_WEIGHT.ColumnName = BtwHdr.c40_PACKAGE_NET_WEIGHT; _dtNalepky.Columns.Add(c40_PACKAGE_NET_WEIGHT);
            DataColumn c41_EXPLOSIVE_NET_WEIGHT = new DataColumn(); c41_EXPLOSIVE_NET_WEIGHT.DataType = System.Type.GetType("System.String"); c41_EXPLOSIVE_NET_WEIGHT.ColumnName = BtwHdr.c41_EXPLOSIVE_NET_WEIGHT; _dtNalepky.Columns.Add(c41_EXPLOSIVE_NET_WEIGHT);
            DataColumn c42_UN_NO = new DataColumn(); c42_UN_NO.DataType = System.Type.GetType("System.String"); c42_UN_NO.ColumnName = BtwHdr.c42_UN_NO; _dtNalepky.Columns.Add(c42_UN_NO);
            DataColumn c43_ADR_RID_CLASS = new DataColumn(); c43_ADR_RID_CLASS.DataType = System.Type.GetType("System.String"); c43_ADR_RID_CLASS.ColumnName = BtwHdr.c43_ADR_RID_CLASS; _dtNalepky.Columns.Add(c43_ADR_RID_CLASS);
            DataColumn c44_DETONATOR_CLASS = new DataColumn(); c44_DETONATOR_CLASS.DataType = System.Type.GetType("System.String"); c44_DETONATOR_CLASS.ColumnName = BtwHdr.c44_DETONATOR_CLASS; _dtNalepky.Columns.Add(c44_DETONATOR_CLASS);
            DataColumn c45_DELAY_LEVEL = new DataColumn(); c45_DELAY_LEVEL.DataType = System.Type.GetType("System.String"); c45_DELAY_LEVEL.ColumnName = BtwHdr.c45_DELAY_LEVEL; _dtNalepky.Columns.Add(c45_DELAY_LEVEL);
            DataColumn c46_CAVITY_MATERIAL = new DataColumn(); c46_CAVITY_MATERIAL.DataType = System.Type.GetType("System.String"); c46_CAVITY_MATERIAL.ColumnName = BtwHdr.c46_CAVITY_MATERIAL; _dtNalepky.Columns.Add(c46_CAVITY_MATERIAL);
            DataColumn c47_IGNITION_CUE_MJ = new DataColumn(); c47_IGNITION_CUE_MJ.DataType = System.Type.GetType("System.String"); c47_IGNITION_CUE_MJ.ColumnName = BtwHdr.c47_IGNITION_CUE_MJ; _dtNalepky.Columns.Add(c47_IGNITION_CUE_MJ);
            DataColumn c48_IGNITION_CUE_J = new DataColumn(); c48_IGNITION_CUE_J.DataType = System.Type.GetType("System.String"); c48_IGNITION_CUE_J.ColumnName = BtwHdr.c48_IGNITION_CUE_J; _dtNalepky.Columns.Add(c48_IGNITION_CUE_J);
            DataColumn c49_SAFETY_STREAM = new DataColumn(); c49_SAFETY_STREAM.DataType = System.Type.GetType("System.String"); c49_SAFETY_STREAM.ColumnName = BtwHdr.c49_SAFETY_STREAM; _dtNalepky.Columns.Add(c49_SAFETY_STREAM);
            DataColumn c50_SAFETY_CUE_MJ = new DataColumn(); c50_SAFETY_CUE_MJ.DataType = System.Type.GetType("System.String"); c50_SAFETY_CUE_MJ.ColumnName = BtwHdr.c50_SAFETY_CUE_MJ; _dtNalepky.Columns.Add(c50_SAFETY_CUE_MJ);
            DataColumn c51_SAFETY_CUE_J = new DataColumn(); c51_SAFETY_CUE_J.DataType = System.Type.GetType("System.String"); c51_SAFETY_CUE_J.ColumnName = BtwHdr.c51_SAFETY_CUE_J; _dtNalepky.Columns.Add(c51_SAFETY_CUE_J);
            DataColumn c52_SUPPLY_WIRE_MATERIAL = new DataColumn(); c52_SUPPLY_WIRE_MATERIAL.DataType = System.Type.GetType("System.String"); c52_SUPPLY_WIRE_MATERIAL.ColumnName = BtwHdr.c52_SUPPLY_WIRE_MATERIAL; _dtNalepky.Columns.Add(c52_SUPPLY_WIRE_MATERIAL);
            DataColumn c53_WIRE_ISULATOR_MATERIAL = new DataColumn(); c53_WIRE_ISULATOR_MATERIAL.DataType = System.Type.GetType("System.String"); c53_WIRE_ISULATOR_MATERIAL.ColumnName = BtwHdr.c53_WIRE_ISULATOR_MATERIAL; _dtNalepky.Columns.Add(c53_WIRE_ISULATOR_MATERIAL);
            DataColumn c54_WIRE_LENGTH = new DataColumn(); c54_WIRE_LENGTH.DataType = System.Type.GetType("System.String"); c54_WIRE_LENGTH.ColumnName = BtwHdr.c54_WIRE_LENGTH; _dtNalepky.Columns.Add(c54_WIRE_LENGTH);
            DataColumn c55_WIRE_LENGTH_TRACE = new DataColumn(); c55_WIRE_LENGTH_TRACE.DataType = System.Type.GetType("System.String"); c55_WIRE_LENGTH_TRACE.ColumnName = BtwHdr.c55_WIRE_LENGTH_TRACE; _dtNalepky.Columns.Add(c55_WIRE_LENGTH_TRACE);
            DataColumn c56_CONSUMPTION_DATE = new DataColumn(); c56_CONSUMPTION_DATE.DataType = System.Type.GetType("System.String"); c56_CONSUMPTION_DATE.ColumnName = BtwHdr.c56_CONSUMPTION_DATE; _dtNalepky.Columns.Add(c56_CONSUMPTION_DATE);
            DataColumn c57_NOMINAL_TIME_EA = new DataColumn(); c57_NOMINAL_TIME_EA.DataType = System.Type.GetType("System.String"); c57_NOMINAL_TIME_EA.ColumnName = BtwHdr.c57_NOMINAL_TIME_EA; _dtNalepky.Columns.Add(c57_NOMINAL_TIME_EA);
            DataColumn c58_NOMINAL_TIME_NA = new DataColumn(); c58_NOMINAL_TIME_NA.DataType = System.Type.GetType("System.String"); c58_NOMINAL_TIME_NA.ColumnName = BtwHdr.c58_NOMINAL_TIME_NA; _dtNalepky.Columns.Add(c58_NOMINAL_TIME_NA);
            DataColumn c59_ST_LENGTH = new DataColumn(); c59_ST_LENGTH.DataType = System.Type.GetType("System.String"); c59_ST_LENGTH.ColumnName = BtwHdr.c59_ST_LENGTH; _dtNalepky.Columns.Add(c59_ST_LENGTH);
            DataColumn c60_DELAY_STEP = new DataColumn(); c60_DELAY_STEP.DataType = System.Type.GetType("System.String"); c60_DELAY_STEP.ColumnName = BtwHdr.c60_DELAY_STEP; _dtNalepky.Columns.Add(c60_DELAY_STEP);
            DataColumn c61_BATCH_NO = new DataColumn(); c61_BATCH_NO.DataType = System.Type.GetType("System.String"); c61_BATCH_NO.ColumnName = BtwHdr.c61_BATCH_NO; _dtNalepky.Columns.Add(c61_BATCH_NO);
            DataColumn c62_TESTING_STATION_NUMBER = new DataColumn(); c62_TESTING_STATION_NUMBER.DataType = System.Type.GetType("System.String"); c62_TESTING_STATION_NUMBER.ColumnName = BtwHdr.c62_TESTING_STATION_NUMBER; _dtNalepky.Columns.Add(c62_TESTING_STATION_NUMBER);
            DataColumn c63_TESTING_PROTOCOL_NUMBER = new DataColumn(); c63_TESTING_PROTOCOL_NUMBER.DataType = System.Type.GetType("System.String"); c63_TESTING_PROTOCOL_NUMBER.ColumnName = BtwHdr.c63_TESTING_PROTOCOL_NUMBER; _dtNalepky.Columns.Add(c63_TESTING_PROTOCOL_NUMBER);
            DataColumn c64_CUSTOMER_AUTHORIZE_DECREE = new DataColumn(); c64_CUSTOMER_AUTHORIZE_DECREE.DataType = System.Type.GetType("System.String"); c64_CUSTOMER_AUTHORIZE_DECREE.ColumnName = BtwHdr.c64_CUSTOMER_AUTHORIZE_DECREE; _dtNalepky.Columns.Add(c64_CUSTOMER_AUTHORIZE_DECREE);
            DataColumn c65_ANOTHER_NOTICE = new DataColumn(); c65_ANOTHER_NOTICE.DataType = System.Type.GetType("System.String"); c65_ANOTHER_NOTICE.ColumnName = BtwHdr.c65_ANOTHER_NOTICE; _dtNalepky.Columns.Add(c65_ANOTHER_NOTICE);
            DataColumn c66_NOTE = new DataColumn(); c66_NOTE.DataType = System.Type.GetType("System.String"); c66_NOTE.ColumnName = BtwHdr.c66_NOTE; _dtNalepky.Columns.Add(c66_NOTE);
            DataColumn c67_OBJVERSION = new DataColumn(); c67_OBJVERSION.DataType = System.Type.GetType("System.String"); c67_OBJVERSION.ColumnName = BtwHdr.c67_OBJVERSION; _dtNalepky.Columns.Add(c67_OBJVERSION);
            DataColumn c68_OBJID = new DataColumn(); c68_OBJID.DataType = System.Type.GetType("System.String"); c68_OBJID.ColumnName = BtwHdr.c68_OBJID; _dtNalepky.Columns.Add(c68_OBJID);
            DataColumn c69_DATE = new DataColumn(); c69_DATE.DataType = System.Type.GetType("System.String"); c69_DATE.ColumnName = BtwHdr.c69_DATE; _dtNalepky.Columns.Add(c69_DATE);
            DataColumn c70_EXPIRATION = new DataColumn(); c70_EXPIRATION.DataType = System.Type.GetType("System.String"); c70_EXPIRATION.ColumnName = BtwHdr.c70_EXPIRATION; _dtNalepky.Columns.Add(c70_EXPIRATION);
            DataColumn c71_TOTAL_RESIST = new DataColumn(); c71_TOTAL_RESIST.DataType = System.Type.GetType("System.String"); c71_TOTAL_RESIST.ColumnName = BtwHdr.c71_TOTAL_RESIST; _dtNalepky.Columns.Add(c71_TOTAL_RESIST);
            DataColumn c72_BRIDGE_RESIST = new DataColumn(); c72_BRIDGE_RESIST.DataType = System.Type.GetType("System.String"); c72_BRIDGE_RESIST.ColumnName = BtwHdr.c72_BRIDGE_RESIST; _dtNalepky.Columns.Add(c72_BRIDGE_RESIST);
            DataColumn c73_DATE_CODE = new DataColumn(); c73_DATE_CODE.DataType = System.Type.GetType("System.String"); c73_DATE_CODE.ColumnName = BtwHdr.c73_DATE_CODE; _dtNalepky.Columns.Add(c73_DATE_CODE);
            DataColumn c74_LOT_NUMBER = new DataColumn(); c74_LOT_NUMBER.DataType = System.Type.GetType("System.String"); c74_LOT_NUMBER.ColumnName = BtwHdr.c74_LOT_NUMBER; _dtNalepky.Columns.Add(c74_LOT_NUMBER);
            DataColumn c75_STATION_NUM = new DataColumn(); c75_STATION_NUM.DataType = System.Type.GetType("System.String"); c75_STATION_NUM.ColumnName = BtwHdr.c75_STATION_NUM; _dtNalepky.Columns.Add(c75_STATION_NUM);
            DataColumn c76_WORKERS = new DataColumn(); c76_WORKERS.DataType = System.Type.GetType("System.String"); c76_WORKERS.ColumnName = BtwHdr.c76_WORKERS; _dtNalepky.Columns.Add(c76_WORKERS);
            DataColumn c77_WORKER2 = new DataColumn(); c77_WORKER2.DataType = System.Type.GetType("System.String"); c77_WORKER2.ColumnName = BtwHdr.c77_WORKER2; _dtNalepky.Columns.Add(c77_WORKER2);
            DataColumn c78_DUPLICATE = new DataColumn(); c78_DUPLICATE.DataType = System.Type.GetType("System.String"); c78_DUPLICATE.ColumnName = BtwHdr.c78_DUPLICATE; _dtNalepky.Columns.Add(c78_DUPLICATE);
            DataColumn c79_PIECE_NUM = new DataColumn(); c79_PIECE_NUM.DataType = System.Type.GetType("System.String"); c79_PIECE_NUM.ColumnName = BtwHdr.c79_PIECE_NUM; _dtNalepky.Columns.Add(c79_PIECE_NUM);
            DataColumn c80_DELAY_FRACTION = new DataColumn(); c80_DELAY_FRACTION.DataType = System.Type.GetType("System.String"); c80_DELAY_FRACTION.ColumnName = BtwHdr.c80_DELAY_FRACTION; _dtNalepky.Columns.Add(c80_DELAY_FRACTION);
            DataColumn c81_GENERAL = new DataColumn(); c81_GENERAL.DataType = System.Type.GetType("System.String"); c81_GENERAL.ColumnName = BtwHdr.c81_GENERAL; _dtNalepky.Columns.Add(c81_GENERAL);
            DataColumn c82_GENERAL_1 = new DataColumn(); c82_GENERAL_1.DataType = System.Type.GetType("System.String"); c82_GENERAL_1.ColumnName = BtwHdr.c82_GENERAL_1; _dtNalepky.Columns.Add(c82_GENERAL_1);
            DataColumn c83_BOX_PCS = new DataColumn(); c83_BOX_PCS.DataType = System.Type.GetType("System.String"); c83_BOX_PCS.ColumnName = BtwHdr.c83_BOX_PCS; _dtNalepky.Columns.Add(c83_BOX_PCS);
            DataColumn c84_MANUAL = new DataColumn(); c84_MANUAL.DataType = System.Type.GetType("System.String"); c84_MANUAL.ColumnName = BtwHdr.c84_MANUAL; _dtNalepky.Columns.Add(c84_MANUAL);
            DataColumn c85_PRODUCER = new DataColumn(); c85_PRODUCER.DataType = System.Type.GetType("System.String"); c85_PRODUCER.ColumnName = BtwHdr.c85_PRODUCER; _dtNalepky.Columns.Add(c85_PRODUCER);
            DataColumn c86_DMX_PACKAGE_QTY = new DataColumn(); c86_DMX_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c86_DMX_PACKAGE_QTY.ColumnName = BtwHdr.c86_DMX_PACKAGE_QTY; _dtNalepky.Columns.Add(c86_DMX_PACKAGE_QTY);
            DataColumn c87_DMX_DELAY_LEVEL = new DataColumn(); c87_DMX_DELAY_LEVEL.DataType = System.Type.GetType("System.String"); c87_DMX_DELAY_LEVEL.ColumnName = BtwHdr.c87_DMX_DELAY_LEVEL; _dtNalepky.Columns.Add(c87_DMX_DELAY_LEVEL);
            DataColumn c88_DMX_DELAY_STEP = new DataColumn(); c88_DMX_DELAY_STEP.DataType = System.Type.GetType("System.String"); c88_DMX_DELAY_STEP.ColumnName = BtwHdr.c88_DMX_DELAY_STEP; _dtNalepky.Columns.Add(c88_DMX_DELAY_STEP);
            DataColumn c89_DMX_NOMINAL_TIME_EA = new DataColumn(); c89_DMX_NOMINAL_TIME_EA.DataType = System.Type.GetType("System.String"); c89_DMX_NOMINAL_TIME_EA.ColumnName = BtwHdr.c89_DMX_NOMINAL_TIME_EA; _dtNalepky.Columns.Add(c89_DMX_NOMINAL_TIME_EA);
            DataColumn c90_DMX_NOMINAL_TIME_NA = new DataColumn(); c90_DMX_NOMINAL_TIME_NA.DataType = System.Type.GetType("System.String"); c90_DMX_NOMINAL_TIME_NA.ColumnName = BtwHdr.c90_DMX_NOMINAL_TIME_NA; _dtNalepky.Columns.Add(c90_DMX_NOMINAL_TIME_NA);
            DataColumn c91_DMX_WIRE_LENGTH = new DataColumn(); c91_DMX_WIRE_LENGTH.DataType = System.Type.GetType("System.String"); c91_DMX_WIRE_LENGTH.ColumnName = BtwHdr.c91_DMX_WIRE_LENGTH; _dtNalepky.Columns.Add(c91_DMX_WIRE_LENGTH);
            DataColumn c92_DMX_ST_LENGTH = new DataColumn(); c92_DMX_ST_LENGTH.DataType = System.Type.GetType("System.String"); c92_DMX_ST_LENGTH.ColumnName = BtwHdr.c92_DMX_ST_LENGTH; _dtNalepky.Columns.Add(c92_DMX_ST_LENGTH);
            DataColumn c93_DMX_DATE = new DataColumn(); c93_DMX_DATE.DataType = System.Type.GetType("System.String"); c93_DMX_DATE.ColumnName = BtwHdr.c93_DMX_DATE; _dtNalepky.Columns.Add(c93_DMX_DATE);
            DataColumn c94_DMX_PACKAGE_NUM2 = new DataColumn(); c94_DMX_PACKAGE_NUM2.DataType = System.Type.GetType("System.String"); c94_DMX_PACKAGE_NUM2.ColumnName = BtwHdr.c94_DMX_PACKAGE_NUM2; _dtNalepky.Columns.Add(c94_DMX_PACKAGE_NUM2);
            DataColumn c95_DMX_SEQ_NUM = new DataColumn(); c95_DMX_SEQ_NUM.DataType = System.Type.GetType("System.String"); c95_DMX_SEQ_NUM.ColumnName = BtwHdr.c95_DMX_SEQ_NUM; _dtNalepky.Columns.Add(c95_DMX_SEQ_NUM);
            DataColumn c96_DMX_PRODUCER = new DataColumn(); c96_DMX_PRODUCER.DataType = System.Type.GetType("System.String"); c96_DMX_PRODUCER.ColumnName = BtwHdr.c96_DMX_PRODUCER; _dtNalepky.Columns.Add(c96_DMX_PRODUCER);
            DataColumn c97_AUTO_LABEL = new DataColumn(); c97_AUTO_LABEL.DataType = System.Type.GetType("System.String"); c97_AUTO_LABEL.ColumnName = BtwHdr.c97_AUTO_LABEL; _dtNalepky.Columns.Add(c97_AUTO_LABEL);
            DataColumn c98_TOTAL_PACKAGE_QTY = new DataColumn(); c98_TOTAL_PACKAGE_QTY.DataType = System.Type.GetType("System.String"); c98_TOTAL_PACKAGE_QTY.ColumnName = BtwHdr.c98_TOTAL_PACKAGE_QTY; _dtNalepky.Columns.Add(c98_TOTAL_PACKAGE_QTY);
            DataColumn c99_UFI_CODE = new DataColumn(); c99_UFI_CODE.DataType = System.Type.GetType("System.String"); c99_UFI_CODE.ColumnName = BtwHdr.c99_UFI_CODE; _dtNalepky.Columns.Add(c99_UFI_CODE);

        }
        private void NalepkyGeneruj()
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

            string sConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;"; ;
            OracleConnection cn;

            string sLine;
            string sWhereOrderKomplet = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.RELEASE_NO + "'";
            string sWhereOrderLineKomplet = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.RELEASE_NO + "'" +
                            " AND  line_item_no = " + ZO.LINE_ITEM_NO;

            string sWhereOrderLineKompletForSalesPart = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.LINE_ITEM_NO + "'" +
                            " AND  line_item_no <= 0";

            string sBatchNo = ZO.ZO_BatchNo;
            string sPACKAGE_QTY_MULTI = string.Empty;
            string sNOMINAL_TIME_NA_MULTI = string.Empty;
            string sDelay_step_multi = string.Empty;
            string sAUTO_LABEL = string.Empty;

            try
            {
                int i = 0;
                string sErr = string.Empty;
                string sSQL_1 = string.Empty;
                Funkce ff = new Funkce();
                // 
                // Najdi úroveň balení pro počty nálepek
                string sSQL_Uroven = "SELECT uroven FROM LabelsKody WHERE kod = '" + _DOC_CLASS + "'";
                string sUroven = ff.MSSQL_GetOneValue(Globals.ConString_MSSQL_Austin_Trc, "SELECT uroven FROM LabelsKody WHERE kod = '" + _DOC_CLASS + "'");
#pragma warning disable CS0219 // The variable 'iUroven' is assigned but its value is never used
                int iUroven = 0;
#pragma warning restore CS0219 // The variable 'iUroven' is assigned but its value is never used
                switch (sUroven)
                {
                    case "1":
                        iUroven = 1;
                        break;
                    case "2":
                        iUroven = 2;
                        break;
                    case "3":
                        iUroven = 3;
                        break;
                    case "1A":
                        iUroven = 1;
                        break;
                    default:
                        iUroven = 1;
                        break;
                }
                //
                cn = new OracleConnection(sConString);
                cn.Open();
                
                // RECEIVER
                // CUSTOMER_INFO_ADDRESS.ADDRESS2
                sSQL_1 = Globals.SQL_GetIDZakaznika.Replace("xxORDER_NO", ZO.ORDER_NO);
                string sCustomer_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = Globals.SQL_GetIDAddress.Replace("xxORDER_NO", ZO.ORDER_NO);
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
                if (Globals.TNSNAME.Contains("UPG"))
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
                
                //btw.VENDOR_STREET
                sSQL_1 = Globals.SQL_GetVendor_Address;
                string sVENDOR_ADDRESS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //btw.VENDOR_ZIP_CODE
                sSQL_1 = Globals.SQL_GetVendor_ZipCode;
                string sVENDOR_ZIP_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COMPANY = 'AD'
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_COMPANY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COUNTRY
                sSQL_1 = Globals.SQL_Get_Language_Code_FromOrder.Replace("xxORDER_NO", ZO.ORDER_NO);
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
                            " AND  a.order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  a.line_no = '" + ZO.LINE_NO + "'" +
                            " AND  a.rel_no = '" + ZO.RELEASE_NO + "'";

                bool bUsporneBaleni = ff.Oracle_Exists(sConString, sSQL_UsporneBaleni, ref sErr);
                if (bUsporneBaleni)
                {
                    string sSQL_Multi = "SELECT * FROM (SELECT distinct order_no,package_num2,part_no," +
                        "replace(wm_concat(amount) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') PACKAGE_QTY_MULTI," +
                        "replace(wm_concat(nominal_time) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') NOMINAL_TIME_NA_MULTI," +
                        "replace(wm_concat(delay_step_multi) OVER (PARTITION BY package_num2 ORDER BY order_no,package_num2) ,',','+') Delay_step_multi" +
                        " FROM" +
                        "( " +
                        "   SELECT a.order_no,b.package_num2,b.chZO.LINE_ITEM_NO,b.amount,b.part_no," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'N016') as nominal_time," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'L004') as delay_step_multi" +
                        "   FROM  IFSAPP.C_PACKAGE_HEAD a, IFSAPP.C_PACKAGE_DETAIL b " +
                        "   WHERE a.order_no = b.order_no" +
                        "   AND a.line_no = b.line_no" +
                        "   AND a.rel_no = b.rel_no" +
                        "   AND a.line_item_no = b.line_item_no" +
                        "   AND a.unique_pkg = 'FALSE'" +
                        "   AND  a.order_no = '" + ZO.ORDER_NO + "'" +
                        "   AND  a.line_no = '" + ZO.LINE_NO + "'" +
                        "   AND  a.rel_no = '" + ZO.RELEASE_NO + "'" +
                        "   AND  a.line_item_no = '" + ZO.LINE_ITEM_NO + "'" +
                        "   ORDER BY order_no,LINE_ITEM_NO" +
                        "))" +
                        " WHERE part_no = '" + ZO.PART_NO + "'" +
                        " ORDER BY order_no,package_num2";
                    // při úsporném balení vrátí více řádků (kartonů) pro výrobu jedné položky 
                    sPACKAGE_QTY_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "PACKAGE_QTY_MULTI", ref sErr);
                    sNOMINAL_TIME_NA_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "NOMINAL_TIME_NA_MULTI", ref sErr);
                    sDelay_step_multi = ff.Oracle_GetValue(sConString, sSQL_Multi, "delay_step_multi", ref sErr);


                }

                // KARTONY
                // výběr kartonů  ze zakázek z IFS 
                string sSQL_K = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_K = sSQL_K.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_K = sSQL_K.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_K = sSQL_K.Replace("xxTypNalepky", "KTN");
                OracleCommand cmdK = new OracleCommand(sSQL_K, cn);
                OracleDataReader rdK = cmdK.ExecuteReader();
                if (!rdK.HasRows)
                {
                    MessageBox.Show("Nejsou kartony. Jak to ??");
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
                }
                if ((i == 1))
                {
                    sLine = "FILENAME_KTN2 =";
                }
                rdK.Close();
                cmdK.Dispose();


                //  pro Skladacku
                string sSQL_S = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "STN");

                OracleCommand cmdS = new OracleCommand(sSQL_S, cn);
                OracleDataReader rdS = cmdS.ExecuteReader();
                if (!rdS.HasRows)
                {
                    throw new Exception("Nenašel nálepky N-NA-STN !!");
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
                        sAUTO_LABEL = ff.MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'");
                        sAUTO_LABEL = "AUTO_LABEL = " + sAUTO_LABEL;

                    }
                    if (rdS["Title"].GetType() != typeof(DBNull))
                    {
                        sTitle = rdS["Title"].ToString();
                    }
                }
                rdS.Close();
                cmdS.Dispose();

                //  pro Praporek
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
                        //" AND a.key_ref LIKE  '%ORDER_NO=" + ZO.ORDER_NO + "%'" +
                        //" AND a.key_ref LIKE  '%LINE_NO=" + ZO.LINE_NO + "%'" +
                        //" AND a.key_ref LIKE  '%REL_NO=" + ZO.RELEASE_NO + "%'" +
                        " AND a.key_value LIKE  '" + ZO.ORDER_NO + "^" + ZO.LINE_NO + "^" + ZO.RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";


                OracleCommand cmdP = new OracleCommand(sSQL_P, cn);
                OracleDataReader rdP = cmdP.ExecuteReader();
                if (!rdP.HasRows) // NOT !!!!!!!!!!!
                {
                    //throw new Exception("Nenašel nálepky PRO !!");
                    MessageBox.Show("Nenašel nálepky PRO !!");
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
                            }
                            if (sPRO.Contains(".BTW") && bPRO_BTW == true && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
                            }
                            if (sPRO.Contains(".ZPL") && bPRO_BTW == false && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
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
                sSQL_S = Globals.SQL_DOC_CLASS.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxFILE_NAME", _FileNameBTW);
                sErr = ff.Oracle_GetValue3(sConString, sSQL_S, ref sDOC_CLASS, ref sDOC_NO, ref sDOC_REV);


                //CUSTOMER_PO_NO 
                sSQL_1 = "SELECT CUSTOMER_PO_NO FROM IFSAPP.CUSTOMER_ORDER WHERE order_no = '" + ZO.ORDER_NO + "'";
                string sCUSTOMER_PO_NO = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CUSTOMER_PO_NO = " + sCUSTOMER_PO_NO;

                //ORDER_NO
                sLine = "ORDER_NO = " + ZO.ORDER_NO;

                //LINE_NO
                sLine = "LINE_NO = " + ZO.LINE_NO;

                //RELEASE_NO
                sLine = "RELEASE_NO = " + ZO.RELEASE_NO;

                //LINE_ITEM_NO
                sLine = "LINE_ITEM_NO = " + ZO.LINE_ITEM_NO;

                //PACKAGE_NUM1
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM1) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM1 = " + sPACKAGE_NUM1;

                //PACKAGE_NUM2
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2 = " + sPACKAGE_NUM2;

                //CATALOG_NO
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCatalog_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_NO = " + sCatalog_no;

                //CATALOG_DESC
                sSQL_1 = "SELECT  CATALOG_DESC FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCATALOG_DESC = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_DESC = " + sCATALOG_DESC;

                //PACKAGE_QTY
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT  IFSAPP.C_Package_Head_Cfp.Get_Cf$_Package_Capacity(PACKING_INSTRUCTION_ID) as package_qty FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT  AMOUNT FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }
                
                string sPACKAGE_QTY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iPACKAGE_QTY = 0;
                if (Int32.TryParse(sPACKAGE_QTY, out iPACKAGE_QTY)) { } else { MessageBox.Show("Nenašel množství v kartonu !!"); return; };
                sLine = "PACKAGE_QTY = " + sPACKAGE_QTY;

                //SMALL_PACKAGE_QTY 120/4
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT b.quantity FROM IFSAPP.C_PACKAGE_HEAD_TAB a, IFSAPP.PACKING_INSTRUCTION_NODE_TAB b WHERE  a.packing_instruction_id = b.packing_instruction_id AND b.handling_unit_type_id = 'INNER' AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT IFSAPP.C_PACKAGE_HEAD_API.Get_Inner_Package_Count(PACKAGE_TYPE,TEMPLATE_ID) FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }

                string sSEQ_NUM = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iSEQ_NUM = 0;
                if (Int32.TryParse(sSEQ_NUM, out iSEQ_NUM)) { } else { MessageBox.Show("Nenašel počet vnitřních obalů !!"); return; };
                int iSMALL_PACKAGE_QTY = iPACKAGE_QTY / iSEQ_NUM;
                sLine = "SMALL_PACKAGE_QTY = " + iSMALL_PACKAGE_QTY.ToString();

                //PACKAGE_GROSS_WEIGHT
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    // dohledat konkrétní karton
                    sSQL_1 = "SELECT  gross_weight FROM IFSAPP.C_PACKAGE_WEIGHT WHERE " + sWhereOrderLineKomplet;
                    return;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }

                string sPACKAGE_GROSS_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sPACKAGE_GROSS_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_GROSS_WEIGHT));
                sPACKAGE_GROSS_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_GROSS_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_GROSS_WEIGHT);
                sLine = "PACKAGE_GROSS_WEIGHT = " + sPACKAGE_GROSS_WEIGHT;

                //ROUND_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sROUND_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sROUND_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##}", Convert.ToDouble(sROUND_NET_WEIGHT));
                sROUND_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sROUND_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sROUND_NET_WEIGHT);
                sLine = "ROUND_NET_WEIGHT = " + sROUND_NET_WEIGHT;

                //PACKAGE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sPACKAGE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sPACKAGE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:#####.0}", Convert.ToDouble(sPACKAGE_NET_WEIGHT));
                sPACKAGE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_NET_WEIGHT);
                sLine = "PACKAGE_NET_WEIGHT = " + sPACKAGE_NET_WEIGHT;

                //EXPLOSIVE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sEXPLOSIVE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sEXPLOSIVE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:#####.000}", Convert.ToDouble(sEXPLOSIVE_NET_WEIGHT));
                sEXPLOSIVE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sEXPLOSIVE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sEXPLOSIVE_NET_WEIGHT);
                sLine = "EXPLOSIVE_NET_WEIGHT = " + sEXPLOSIVE_NET_WEIGHT;

                //UN_NO
                sSQL_1 = "SELECT  un FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sUN = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "UN_NO = " + sUN;

                //ADR_RID_CLASS
                sSQL_1 = "SELECT class FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sADR_RID_CLASS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ADR_RID_CLASS = " + sADR_RID_CLASS;
                //DETONATOR_CLASS
                string sDETONATOR_CLASS = string.Empty;
                sLine = "DETONATOR_CLASS = "; // u DA není
                                              //sw.WriteLine(sLine);

                //CONSUMPTION_DATE
                sSQL_1 = "SELECT part_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sPart_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT consumption_month FROM IFSAPP.a0000_polozky_char WHERE part_no = '" + sPart_no + "' AND contract = 'AD'";
                string sConsumption_month = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = "SELECT to_char(add_months(wanted_delivery_date," + sConsumption_month + "),'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCONSUMPTION_DATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONSUMPTION_DATE = " + sCONSUMPTION_DATE;

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

                //DELAY_STEP = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L004'";
                string sDELAY_STEP = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_STEP = " + sDELAY_STEP; // u DA není

                //DELAY_STEP_MULTI  
                sLine = "DELAY_STEP_MULTI = "; // u DA není

                // CAVITY_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L003'";
                string sCAVITY_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CAVITY_MATERIAL  = " + sCAVITY_MATERIAL;

                //SUPPLY_WIRE_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L008'";
                string sSUPPLY_WIRE_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SUPPLY_WIRE_MATERIAL = " + sSUPPLY_WIRE_MATERIAL;

                //WIRE_ISULATOR_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L009'";
                string sWIRE_ISULATOR_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_ISULATOR_MATERIAL = " + sWIRE_ISULATOR_MATERIAL;

                //WIRE_LENGTH 
                sSQL_1 = "SELECT b.characteristic_value_desc FROM IFSAPP.INV_PART_DISCRETE_CHAR a,IFSAPP.discrete_charac_value b" +
                    " WHERE a.characteristic_code = b.characteristic_code AND a.attr_value = b.characteristic_value" + 
                    " AND  a.part_no = '" + sPart_no + "' AND a.contract = 'AD'" +
                    " AND a.characteristic_code = 'E005'";
                string sWIRE_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_LENGTH = " + sWIRE_LENGTH;

                //NOMINAL_TIME_EA 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sNOMINAL_TIME_EA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_EA = " + sNOMINAL_TIME_EA;

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

                //TESTING_PROTOCOL_NUMBER
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE011'";
                string sTESTING_PROTOCOL_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_PROTOCOL_NUMBER = " + sTESTING_PROTOCOL_NUMBER;

                //CUSTOMER_AUTHORIZE_DECREE
                sSQL_1 = "SELECT dt.output_type FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + ZO.ORDER_NO + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sOutput_type = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT dt.note_text FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + ZO.ORDER_NO + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sNote_text = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                string sCUSTOMER_AUTHORIZE_DECREE = string.Empty;
                switch (sOutput_type)
                {
                    case "LAB01":
                        sLine = "CUSTOMER_AUTHORIZE_DECREE = " + sNote_text;
                        sCUSTOMER_AUTHORIZE_DECREE = sNote_text;
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
                //GENERAL
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GENE'";
                string sGENERAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL = " + sGENERAL;
                //GENERAL_1
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GEN01'";
                string sGENERAL_1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL_1 = " + sGENERAL_1;
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
                sSQL_1 = "SELECT to_char(wanted_delivery_date,'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sDATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DATE = " + sDATE;

                //EXPIRATION
                sLine = "EXPIRATION = " + sCONSUMPTION_DATE;

                //PACKAGE_NUM2BEG
                sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2BEG = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2BEG = " + sPACKAGE_NUM2BEG;

                //PACKAGE_NUM2END
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2END = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2END = " + sPACKAGE_NUM2END;

                //SEQ_TOTAL - počet skládaček sSMALL_PACKAGE_QTY
                sLine = "SEQ_TOTAL = " + iSEQ_NUM.ToString();

                //BOX_PCS
                sLine = "BOX_PCS = " + iSMALL_PACKAGE_QTY.ToString();

                //TOTAL_PCS
                sSQL_1 = "SELECT SUM(amount) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sTOTAL_PCS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TOTAL_PCS = " + sTOTAL_PCS;

                //UFI_CODE
                sSQL_1 = "select ch.CF$_UFI_CODE_A1149 from IFSAPP.A1135_INVPART_CHAR_ALL_CFV ch where ch.CONTRACT = 'AD' and ch.PART_NO = '" + sPart_no + "' AND ch.CHARACTERISTIC_CODE = 'W001'";
                string sUFI_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //sLine = "UFI_CODE = " + sUFI_CODE;

                //PRODUCER
                sLine = "PRODUCER = " + sPRODUCER;
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

                //DMX_PACKAGE_QTY
                sLine = sDMX_TEXT_PACKAGE_QTY + sGS_DMX_PACKAGE_QTY + sPACKAGE_QTY;
                //

                //DMX_DELAY_LEVEL
                sLine = sDMX_TEXT_DELAY_LEVEL + sGS_DMX_DELAY_LEVEL + sDELAY_LEVEL;

                //DMX_DEALY_STEP ;
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sDMX_DEALY_STEP_Value = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = sDMX_TEXT_DELAY_STEP + sGS_DMX_DELAY_STEP + sDMX_DEALY_STEP_Value;

                //DMX_NOMINAL_TIME_EA
                sLine = sDMX_TEXT_NOMINAL_TIME_EA + sGS_DMX_NOMINAL_TIME_EA + sNOMINAL_TIME_EA;

                //DMX_NOMINAL_TIME_NA
                sLine = sDMX_TEXT_NOMINAL_TIME_NA + sGS_DMX_NOMINAL_TIME_NA + sNOMINAL_TIME_NA;

                //DMX_ST_LENGTH
                sLine = sDMX_TEXT_ST_LENGTH + sGS_DMX_ST_LENGTH + sST_LENGTH;

                //DMX_DATE
                sLine = sDMX_TEXT_DATE + sGS_DMX_DATE + sDATE;

                //DMX_PACKAGE_NUM2
                sLine = sDMX_TEXT_PACKAGE_NUM2 + sGS_DMX_PACKAGE_NUM2 + sPACKAGE_NUM2;

                //DMX_SEQ_NUM
                string sDMX_SEQ_NUM = iSEQ_NUM.ToString();
                sLine = sDMX_TEXT_SEQ_NUM + sGS_DMX_SEQ_NUM + sDMX_SEQ_NUM;

                //DMX_WIRE_LENGTH
                string sDMX_WIRE_LENGTH = sWIRE_LENGTH;
                sLine = sDMX_TEXT_WIRE_LENGTH + sGS_DMX_WIRE_LENGTH + sDMX_WIRE_LENGTH;

                //AUTO_LABEL - nové aplikační identifikátory (je to v SL), podle nich si TRC vybírá praporek
                //sw.WriteLine(sAUTO_LABEL);
                //
                //if (checkBoxDETONATOR_APLIC.Checked) // doplněno 22.9.2014 Senkerik
                //{
                //    sLine = "DETONATOR_APPLIC = NO";
                //    sw.WriteLine(sLine);
                //}
                // datatable
                int iBeg = 0;
                int iEnd = 0;
                int iPocetKartonu = 0;
                int iPocetSkladacek = 0;
                Int32.TryParse(sPACKAGE_NUM2BEG, out iBeg);
                Int32.TryParse(sPACKAGE_NUM2END, out iEnd);
                int iOpakovani = (iUroven==1?1: int.Parse(sSEQ_NUM));
                for (int ii = iBeg; ii <= iEnd; ii++)
                {
                    for (int iU = 1; iU <= iOpakovani; iU++)
                    {

                        DataRow dr = _dtNalepky.NewRow();

                        dr["PACKAGE_ID"] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString();
                        dr["SUBPACKAGE_ID"] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString() + "-" + iU.ToString();


                        dr[BtwHdr.c01_DOC_CLASS] = _DOC_CLASS;
                        dr[BtwHdr.c02_TITLE] = sTitle;
                        dr[BtwHdr.c03_DOC_NO] = sDOC_NO;
                        dr[BtwHdr.c04_DOC_REV] = sDOC_REV;
                        dr[BtwHdr.c05_FILE_NAME] = _FileNameBTW;
                        dr[BtwHdr.c06_CUSTOMER_PO_NO] = sCUSTOMER_PO_NO;
                        dr[BtwHdr.c07_ORDER_NO] = ZO.ORDER_NO;
                        dr[BtwHdr.c08_LINE_NO] = ZO.LINE_NO;
                        dr[BtwHdr.c09_RELEASE_NO] = ZO.RELEASE_NO;
                        dr[BtwHdr.c10_LINE_ITEM_NO] = ZO.LINE_ITEM_NO;
                        dr[BtwHdr.c11_PACKAGE_NUM1] = sPACKAGE_NUM1;
                        dr[BtwHdr.c12_SEQ_NUM] = iU.ToString();
                        dr[BtwHdr.c13_PACKAGE_NUM2] = ii.ToString(); //sPACKAGE_NUM2;
                        dr[BtwHdr.c14_CATALOG_NO] = sCatalog_no;
                        dr[BtwHdr.c15_CATALOG_DESC] = sCATALOG_DESC;
                        dr[BtwHdr.c16_VENDOR_COMPANY] = sVENDOR_COMPANY;
                        dr[BtwHdr.c17_VENDOR_NAME] = sVENDOR_NAME;
                        dr[BtwHdr.c18_VENDOR_ADDRESS_ID] = sVENDOR_ADDRESS_ID;
                        dr[BtwHdr.c19_VENDOR_STREET] = sVENDOR_ADDRESS;
                        dr[BtwHdr.c20_VENDOR_CITY] = sVENDOR_CITY;
                        dr[BtwHdr.c21_VENDOR_ZIP_CODE] = sVENDOR_ZIP_CODE;
                        dr[BtwHdr.c22_VENDOR_COUNTRY] = sVENDOR_COUNTRY;
                        dr[BtwHdr.c23_RECEIVER_COMPANY_NAME2] = sReceiverCompanyName;
                        dr[BtwHdr.c24_RECEIVER_ADDRESS1] = sReceiverAddress1;
                        dr[BtwHdr.c25_RECEIVER_ADDRESS2] = sReceiverAddress2;
                        dr[BtwHdr.c26_RECEIVER_ZIP_CODE] = sReceiverZIP_CODE;
                        dr[BtwHdr.c27_RECEIVER_CITY] = sReceiverCity;
                        dr[BtwHdr.c28_RECEIVER_COUNTRY] = sReceiverCountry;
                        dr[BtwHdr.c29_RECEIVER_NAME] = sReceiverContactName;
                        dr[BtwHdr.c30_RECEIVER_PHONE] = sReceiverContactPhone;
                        dr[BtwHdr.c31_RECEIVER_FAX] = sReceiverContactFax;
                        dr[BtwHdr.c32_RECEIVER_MOBILE] = sReceiverContactMobile;
                        dr[BtwHdr.c33_RECEIVER_EMAIL] = sReceiverContactEmail;
                        dr[BtwHdr.c34_UNIQUE_PACKAGE_ID] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString() ;
                        dr[BtwHdr.c35_UNIQUE_SUBPACKAGE_ID] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString() + "-" + iU.ToString();
                        dr[BtwHdr.c36_PACKAGE_QTY] = sPACKAGE_QTY;
                        dr[BtwHdr.c37_SMALL_PACKAGE_QTY] = iSMALL_PACKAGE_QTY.ToString();
                        dr[BtwHdr.c38_PACKAGE_GROSS_WEIGHT] = sPACKAGE_GROSS_WEIGHT;
                        dr[BtwHdr.c40_PACKAGE_NET_WEIGHT] = sPACKAGE_NET_WEIGHT;
                        dr[BtwHdr.c41_EXPLOSIVE_NET_WEIGHT] = sEXPLOSIVE_NET_WEIGHT;
                        dr[BtwHdr.c42_UN_NO] = sUN;
                        dr[BtwHdr.c43_ADR_RID_CLASS] = sADR_RID_CLASS;
                        dr[BtwHdr.c44_DETONATOR_CLASS] = sDETONATOR_CLASS;
                        dr[BtwHdr.c45_DELAY_LEVEL] = sDELAY_LEVEL;
                        dr[BtwHdr.c46_CAVITY_MATERIAL] = sCAVITY_MATERIAL;
                        dr[BtwHdr.c47_IGNITION_CUE_MJ] = string.Empty;
                        dr[BtwHdr.c48_IGNITION_CUE_J] = string.Empty;
                        dr[BtwHdr.c49_SAFETY_STREAM] = string.Empty;
                        dr[BtwHdr.c50_SAFETY_CUE_MJ] = string.Empty;
                        dr[BtwHdr.c51_SAFETY_CUE_J] = string.Empty;
                        dr[BtwHdr.c52_SUPPLY_WIRE_MATERIAL] = sSUPPLY_WIRE_MATERIAL;
                        dr[BtwHdr.c53_WIRE_ISULATOR_MATERIAL] = sWIRE_ISULATOR_MATERIAL;
                        dr[BtwHdr.c54_WIRE_LENGTH] = sWIRE_LENGTH;
                        dr[BtwHdr.c55_WIRE_LENGTH_TRACE] = string.Empty;
                        dr[BtwHdr.c56_CONSUMPTION_DATE] = sCONSUMPTION_DATE;
                        dr[BtwHdr.c57_NOMINAL_TIME_EA] = sNOMINAL_TIME_EA;
                        dr[BtwHdr.c58_NOMINAL_TIME_NA] = sNOMINAL_TIME_NA;
                        dr[BtwHdr.c59_ST_LENGTH] = sST_LENGTH;
                        dr[BtwHdr.c60_DELAY_STEP] = sDELAY_STEP;
                        dr[BtwHdr.c61_BATCH_NO] = sBatchNo;
                        dr[BtwHdr.c62_TESTING_STATION_NUMBER] = sTESTING_STATION_NUMBER;
                        dr[BtwHdr.c63_TESTING_PROTOCOL_NUMBER] = sTESTING_PROTOCOL_NUMBER;
                        dr[BtwHdr.c64_CUSTOMER_AUTHORIZE_DECREE] = sCUSTOMER_AUTHORIZE_DECREE;
                        dr[BtwHdr.c65_ANOTHER_NOTICE] = sNote_text;
                        dr[BtwHdr.c66_NOTE] = sNote_text;
                        dr[BtwHdr.c67_OBJVERSION] = string.Empty; // oracle sloupec
                        dr[BtwHdr.c68_OBJID] = string.Empty; // oracle sloupec 
                        dr[BtwHdr.c69_DATE] = sDATE;
                        dr[BtwHdr.c70_EXPIRATION] = sCONSUMPTION_DATE;
                        dr[BtwHdr.c71_TOTAL_RESIST] = string.Empty;
                        dr[BtwHdr.c72_BRIDGE_RESIST] = string.Empty;
                        dr[BtwHdr.c73_DATE_CODE] = string.Empty;
                        dr[BtwHdr.c74_LOT_NUMBER] = string.Empty;
                        dr[BtwHdr.c75_STATION_NUM] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c76_WORKERS] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c77_WORKER2] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c78_DUPLICATE] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c79_PIECE_NUM] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c80_DELAY_FRACTION] = string.Empty; // není v c_labels_ad;
                        dr[BtwHdr.c81_GENERAL] = sGENERAL;
                        dr[BtwHdr.c82_GENERAL_1] = sGENERAL_1;
                        dr[BtwHdr.c83_BOX_PCS] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c84_MANUAL] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c85_PRODUCER] = sPRODUCER;
                        dr[BtwHdr.c86_DMX_PACKAGE_QTY] = sPACKAGE_QTY;
                        dr[BtwHdr.c87_DMX_DELAY_LEVEL] = sDELAY_LEVEL;
                        dr[BtwHdr.c88_DMX_DELAY_STEP] = sDELAY_STEP;
                        dr[BtwHdr.c89_DMX_NOMINAL_TIME_EA] = sNOMINAL_TIME_EA;
                        dr[BtwHdr.c90_DMX_NOMINAL_TIME_NA] = sNOMINAL_TIME_NA;
                        dr[BtwHdr.c91_DMX_WIRE_LENGTH] = sWIRE_LENGTH;
                        dr[BtwHdr.c92_DMX_ST_LENGTH] = sST_LENGTH;
                        dr[BtwHdr.c93_DMX_DATE] = sDATE;
                        dr[BtwHdr.c94_DMX_PACKAGE_NUM2] = sPACKAGE_NUM2;
                        dr[BtwHdr.c95_DMX_SEQ_NUM] = sSEQ_NUM;
                        dr[BtwHdr.c96_DMX_PRODUCER] = sPRODUCER;
                        dr[BtwHdr.c97_AUTO_LABEL] = sAUTO_LABEL;
                        dr[BtwHdr.c98_TOTAL_PACKAGE_QTY] = sTOTAL_PCS;
                        dr[BtwHdr.c99_UFI_CODE] = sTOTAL_PCS;


                        _dtNalepky.Rows.Add(dr);
                        iPocetSkladacek++;
                    }
                    iPocetKartonu++;
                }
                DataRow dr1 = _dtNalepky.NewRow();
                dr1["PACKAGE_ID"] = "Počet:" + iPocetKartonu.ToString();
                dr1["SUBPACKAGE_ID"] = "Počet:" + iPocetSkladacek.ToString();
                _dtNalepky.Rows.Add(dr1);

                dataGridViewLabels.DataSource = _dtNalepky;
                dataGridViewLabels.Refresh();
                for (int jj = 1; jj <= iSEQ_NUM; jj++)
                {
                }
                //

                //Napln_dataGridViewSoubory();
                cn.Close();
                //sw.Dispose();
                //sw.Close();
                //MessageBox.Show("Soubor je vytvořen.");
                //KopirujNalepky();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }
            this.Cursor = Cursors.Default;
        }

        private void dataGridViewBTW_DoubleClick(object sender, EventArgs e)
        {
            if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
            {
                NalepkyGeneruj_IFS9();
            }
            else
            {
                NalepkyGeneruj();
            }
        }

        private void dataGridViewLabels_DoubleClick(object sender, EventArgs e)
        {
            //pictureBoxLabelPreview.Image = null; pictureBoxLabelPreview.Refresh();
            NalepkaNahled(false, dataGridViewLabels.CurrentRow.Index);
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void toolStripMenuItemLabels_Preview_Click(object sender, EventArgs e)
        {
            NalepkaNahled(false, dataGridViewLabels.CurrentRow.Index);

        }

        private void toolStripMenuItemLabels_Print_Click(object sender, EventArgs e)
        {
            if (comboBoxPrinters.SelectedIndex == -1)
            {
                MessageBox.Show("Není vybrána tiskárna");
                return;

            }
            int i = 0;
            foreach (DataGridViewRow r in dataGridViewLabels.SelectedRows)
            {
                //dataGridViewLabels.CurrentRow.Index
                NalepkaNahled(true,r.Index);
                i++;
                r.Selected = false;
            }
            MessageBox.Show("Vytištěno nálepek :  "  + i.ToString(),"Tisk",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void frmNalepkyTisk_FormClosed(object sender, FormClosedEventArgs e)
        {
            Kill_Bartender();
            string supportedExtensions = "*.jpg,*.gif,*.png,*.bmp,*.jpe,*.jpeg";
            foreach (string imageFile in Directory.GetFiles(Globals.FOLDER_BTW_EXPORT, 
                "*.*", SearchOption.AllDirectories).Where(s => supportedExtensions.Contains(Path.GetExtension(s).ToLower())))
            {
                try
                {
                    try
                    {
                        //pictureBoxLabelPreview.Image = null;
                        //pictureBoxLabelPreview.Image.Dispose();

                    }
                    catch (Exception exe1)
                    {
                        //continue;
                    }
                    try
                    {
                        this._Bitmap = null;
                        this._Bitmap.Dispose();
                    }
                    catch (Exception exe2)
                    {
                        //continue;
                    }
                    try
                    {
                        //pictureBoxLabelPreview.Image = null;
                        //pictureBoxLabelPreview.Dispose();
                    }
                    catch (Exception exe)
                    {
                       // continue;
                    }

                    File.Delete(imageFile);
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
        }
        private void Kill_Bartender()
        {
            try
            {
                foreach (System.Diagnostics.Process myProc in System.Diagnostics.Process.GetProcesses())
                {
                    if (myProc.ProcessName == "bartend")
                    {
                        myProc.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonPrintLabel_Click(object sender, EventArgs e)
        {
            //NalepkaNahled(true, dataGridViewLabels.CurrentRow.Index);
            if (comboBoxPrinters.SelectedIndex == -1)
            {
                MessageBox.Show("Není vybrána tiskárna");
                return;

            }
            int i = 0;
            foreach (DataGridViewRow r in dataGridViewLabels.SelectedRows)
            {
                //dataGridViewLabels.CurrentRow.Index
                NalepkaNahled(true, r.Index);
                i++;
                r.Selected = false;
                dataGridViewLabels.Refresh();
            }
            MessageBox.Show("Vytištěno nálepek :  " + i.ToString(), "Tisk", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void LabelPrint()
        {

            if (!File.Exists(_FileNameFolderBTW))
            {
                MessageBox.Show("Soubor :" + Environment.NewLine + _FileNameFolderBTW + Environment.NewLine + " neexistuje !!");
                return;
            }
            if (Path.GetExtension(_FileNameFolderBTW).ToLower() == ".jpg")
            {
                //this._Bitmap = new Bitmap(_FileNameFolderBTW);
                //pictureBoxLabelPreview.Image = (Image)this._Bitmap;
                return;
            }
            if (Path.GetExtension(_FileNameFolderBTW).ToLower() == ".zpl")
            {
                MessageBox.Show("Soubor typu ZPL nejde zobrazit.");
                return;
            }

            string str1 = System.Windows.Forms.Application.StartupPath + (object)'\\';
            BarTender.Format format = (BarTender.Format)null;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                DateTime now = DateTime.Now;
                //
                // Declare a BarTender application variable

                BarTender.Application btApp;
                BarTender.Format btFormat;
                btApp = new BarTender.Application();
                //public static ApplicationClass BTApp = (ApplicationClass)null;
                btApp.Visible = false;

                // Open a BarTender document
                btFormat = btApp.Formats.Open(_FileNameFolderBTW, false, "PDFCreator");

                // Export document to a print preview

                BarTender.BtPrintResult res;
                string sPictureFile = @"\nahled" + DateTime.Now.Millisecond + ".jpg";
                string sPictureFolderFile = Globals.FOLDER_BTW_EXPORT + sPictureFile;
                Globals.JPG_LAST_FOLDER_FILENAME = sPictureFolderFile;


                //S:\KOPY_archiv_vykresova dokumentace\VYKRESY_B A L E N I\aBARTENDER_v901_Nova grafika_data\austin_sablona_1.txt
                //btFormat.Databases.GetDatabase("tisk_data.txt");
                //BtwRow dataBoundItem = (BtwRow)dataGridViewLabels.CurrentRow.DataBoundItem;
                //
                string str2 = btFormat.get_UsedNames(BtNameType.btUsedDatabaseFields, ";");
                //"austin_sablona_1.austin_sablona_1.13_PACKAGE_NUM2;austin_sablona_1.austin_sablona_1.14_CATALOG_NO;austin_sablona_1.austin_sablona_1.36_PACKAGE_QTY;austin_sablona_1.austin_sablona_1.37_SMALL_PACKAGE_QTY;austin_sablona_1.austin_sablona_1.38_PACKAGE_GROSS_WEIGHT;austin_sablona_1.austin_sablona_1.40_PACKAGE_NET_WEIGHT;austin_sablona_1.austin_sablona_1.41_EXPLOSIVE_NET_WEIGHT;austin_sablona_1.austin_sablona_1.42_UN_NO;austin_sablona_1.austin_sablona_1.43_ADR_RID_CLASS;austin_sablona_1.austin_sablona_1.52_SUPPLY_WIRE_MATERIAL;austin_sablona_1.austin_sablona_1.53_WIRE_ISULATOR_MATERIAL;austin_sablona_1.austin_sablona_1.54_WIRE_LENGTH;austin_sablona_1.austin_sablona_1.57_NOMINAL_TIME_EA;austin_sablona_1.austin_sablona_1.61_BATCH_NO;austin_sablona_1.austin_sablona_1.62_TESTING_STATION_NUMBER;austin_sablona_1.austin_sablona_1.63_TESTING_PROTOCOL_NUMBER;austin_sablona_1.austin_sablona_1.69_DATE;austin_sablona_1.austin_sablona_1.70_EXPIRATION;austin_sablona_1.austin_sablona_1.85_PRODUCER"
                string sTable = str2.Substring(0, str2.IndexOf("."));
                string sFileData = Globals.FOLDER_BTW_EXPORT + @"\tisk_data.txt";
                StreamWriter sw = new StreamWriter(sFileData, false, System.Text.Encoding.Default);
                string sLine1 = string.Empty;
                string sLine2 = string.Empty;
                string sPACKAGE_BARCODE_ID = dataGridViewLabels[BtwHdr.c34_UNIQUE_PACKAGE_ID, dataGridViewLabels.CurrentRow.Index].Value.ToString();
                string sSUBPACKAGE_BARCODE_ID = dataGridViewLabels[BtwHdr.c35_UNIQUE_SUBPACKAGE_ID, dataGridViewLabels.CurrentRow.Index].Value.ToString();
                for (int i = 0; i < dataGridViewLabels.ColumnCount; i++)
                {
                    string sSeparator = ((i == dataGridViewLabels.ColumnCount - 1) ? string.Empty : ";");
                    sLine1 = sLine1 + dataGridViewLabels.Columns[i].HeaderText + sSeparator;
                    sLine2 = sLine2 + dataGridViewLabels[i, dataGridViewLabels.CurrentRow.Index].Value.ToString() + sSeparator;
                    string sColumnName = "Textový soubor 1.austin_sablona_1." + dataGridViewLabels.Columns[i].HeaderText;
                    btFormat.set_UsedNamesSampleData(BarTender.BtUsedNamesSampleDataType.btUsedDatabaseFieldsSampleData, sColumnName, dataGridViewLabels[i, dataGridViewLabels.CurrentRow.Index].Value.ToString());
                    //btFormat.SetNamedSubStringValue("KsVKartonu","100");
                }
                sw.WriteLine(sLine1);
                sw.WriteLine(sLine2);
                sw.Close();

                // new with engine
                //Engine btwEngine = new Engine();
                //btwEngine.Start();
                //LabelFormatDocument btwFormat = btwEngine.Documents.Open(_FileNameFolderBTW);
                //btwFormat.DatabaseConnections["Textový soubor 1"].FileName = sFileData;
                //Seagull.BarTender.Print.Messages bteMessages;
                //btwFormat.ExportPrintPreviewToFile(Globals.FOLDER_BTW_EXPORT, sPictureFile, ImageType.BMP,Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300),Color.AliceBlue, OverwriteOptions.DoNotOverwrite, true, true, out bteMessages);
                ////btwFormat.Print();
                //string messageString = "\n\nMessages:";

                //foreach (BarTender.Message message in bteMessages)
                //{
                //    messageString += "\n\n" + message.Message.ToString();
                //}

                //btwEngine.Stop();
                //
                BarTender.Messages btMessages;
                object obj = sTable;
                btFormat.Databases.GetDatabase(obj).TextFile.FileName = sFileData;

                // ***** print
                //btFormat.PrintToFile = true;
                //BarTender.PrinterCodeTemplate bt = btFormat.PrinterCodeTemplate;
                btFormat.Print("Printing Serial", true, -1, out btMessages);
                // *****
                btFormat.Close(BtSaveOptions.btDoNotSaveChanges);
                btFormat = null;
                btApp.Quit(BarTender.BtSaveOptions.btDoNotSaveChanges);
                btApp = null;

            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                int num2 = (int)MessageBox.Show(ex.ToString());
                //Globals.LogError(nameof(btnPreview_Click), "Chyba: " + ex.ToString());
            }
            finally
            {
                if (format != null)
                    format.Close(BtSaveOptions.btDoNotSaveChanges);
                this.Cursor = Cursors.Default;
            }
            //
            //////this.Width = this._Bitmap.Width + 5;
            //////this.Height = this._Bitmap.Height + 40;
            ////pictureBoxLabelPreview.Image = (Image)this._Bitmap;
        }
        private void PrintersGet()
        {
            System.Drawing.Printing.PrinterSettings printerName = new System.Drawing.Printing.PrinterSettings();

            string defaultPrinter;

            defaultPrinter = printerName.PrinterName;

            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                comboBoxPrinters.Items.Add(printer.ToString());
                //MessageBox.Show(printer);
            }
            comboBoxPrinters.SelectedIndex = comboBoxPrinters.Items.IndexOf(defaultPrinter);
            Globals.PRINTER_DEFAULT = defaultPrinter;
        }

        private void comboBoxPrinters_SelectedIndexChanged(object sender, EventArgs e)
        {
            Globals.PRINTER = comboBoxPrinters.Items[comboBoxPrinters.SelectedIndex].ToString();
        }
        private void NalepkyGeneruj_IFS9()
        {
            this.Cursor = Cursors.WaitCursor;
            //
            string sFILENAME_KTN = string.Empty;
            string sFILENAME_KTN2 = string.Empty;
            string sFILENAME_STN = string.Empty;
            string sFILENAME_PRO = string.Empty;
            string sTitle = string.Empty;
            string sPackageBarcodeId = string.Empty;
            //

            MakeDataTable_Nalepky();
            dataGridViewLabels.DataSource = null;

            string sConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456;"; ;
            OracleConnection cn;

            string sLine;
            string sWhereOrderKomplet = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.RELEASE_NO + "'";
            string sWhereOrderLineKomplet = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.RELEASE_NO + "'" +
                            " AND  line_item_no = " + ZO.LINE_ITEM_NO;

            string sWhereOrderLineKompletForSalesPart = " order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  line_no = '" + ZO.LINE_NO + "'" +
                            " AND  rel_no = '" + ZO.LINE_ITEM_NO + "'" +
                            " AND  line_item_no <= 0";

            string sBatchNo = ZO.ZO_BatchNo;
            string sPACKAGE_QTY_MULTI = string.Empty;
            string sNOMINAL_TIME_NA_MULTI = string.Empty;
            string sDelay_step_multi = string.Empty;
            string sAUTO_LABEL = string.Empty;

            try
            {
                int i = 0;
                string sErr = string.Empty;
                string sSQL_1 = string.Empty;
                Funkce ff = new Funkce();
                // 
                // Najdi úroveň balení pro počty nálepek
                string sSQL_Uroven = "SELECT uroven FROM LabelsKody WHERE kod = '" + _DOC_CLASS + "'";
                string sUroven = ff.MSSQL_GetOneValue(Globals.ConString_MSSQL_Austin_Trc, "SELECT uroven FROM LabelsKody WHERE kod = '" + _DOC_CLASS + "'");
#pragma warning disable CS0219 // The variable 'iUroven' is assigned but its value is never used
                int iUroven = 0;
#pragma warning restore CS0219 // The variable 'iUroven' is assigned but its value is never used
                switch (sUroven)
                {
                    case "1":
                        iUroven = 1;
                        break;
                    case "2":
                        iUroven = 2;
                        break;
                    case "3":
                        iUroven = 3;
                        break;
                    case "1A":
                        iUroven = 1;
                        break;
                    default:
                        iUroven = 1;
                        break;
                }
                //
                cn = new OracleConnection(sConString);
                cn.Open();

                // RECEIVER
                // CUSTOMER_INFO_ADDRESS.ADDRESS2
                sSQL_1 = Globals.SQL_GetIDZakaznika.Replace("xxORDER_NO", ZO.ORDER_NO);
                string sCustomer_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = Globals.SQL_GetIDAddress.Replace("xxORDER_NO", ZO.ORDER_NO);
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
                if (Globals.TNSNAME.Contains("UPG"))
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

                //btw.VENDOR_STREET
                sSQL_1 = Globals.SQL_GetVendor_Address;
                string sVENDOR_ADDRESS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //btw.VENDOR_ZIP_CODE
                sSQL_1 = Globals.SQL_GetVendor_ZipCode;
                string sVENDOR_ZIP_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COMPANY = 'AD'
                sSQL_1 = Globals.SQL_GetVendor_Name;
                string sVENDOR_COMPANY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                //
                //btw.VENDOR_COUNTRY
                sSQL_1 = Globals.SQL_Get_Language_Code_FromOrder.Replace("xxORDER_NO", ZO.ORDER_NO);
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
                            " AND  a.order_no = '" + ZO.ORDER_NO + "'" +
                            " AND  a.line_no = '" + ZO.LINE_NO + "'" +
                            " AND  a.rel_no = '" + ZO.RELEASE_NO + "'";
                // sadované?
                //
                bool bUsporneBaleni = ff.Oracle_Exists(sConString, sSQL_UsporneBaleni, ref sErr);
                if (bUsporneBaleni)
                {
                    string sSQL_Multi = "SELECT * FROM (SELECT distinct order_no,package_num2,part_no," +
                        "replace(wm_concat(amount) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') PACKAGE_QTY_MULTI," +
                        "replace(wm_concat(nominal_time) OVER (PARTITION BY order_no,package_num2 ORDER BY order_no,package_num2) ,',','+') NOMINAL_TIME_NA_MULTI," +
                        "replace(wm_concat(delay_step_multi) OVER (PARTITION BY package_num2 ORDER BY order_no,package_num2) ,',','+') Delay_step_multi" +
                        " FROM" +
                        "( " +
                        "   SELECT a.order_no,b.package_num2,b.chZO.LINE_ITEM_NO,b.amount,b.part_no," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'N016') as nominal_time," +
                        "   (SELECT dc.attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR dc where dc.part_no = b.part_no AND dc.contract ='AD' AND dc.CHARACTERISTIC_CODE = 'L004') as delay_step_multi" +
                        "   FROM  IFSAPP.C_PACKAGE_HEAD a, IFSAPP.C_PACKAGE_DETAIL b " +
                        "   WHERE a.order_no = b.order_no" +
                        "   AND a.line_no = b.line_no" +
                        "   AND a.rel_no = b.rel_no" +
                        "   AND a.line_item_no = b.line_item_no" +
                        "   AND a.unique_pkg = 'FALSE'" +
                        "   AND  a.order_no = '" + ZO.ORDER_NO + "'" +
                        "   AND  a.line_no = '" + ZO.LINE_NO + "'" +
                        "   AND  a.rel_no = '" + ZO.RELEASE_NO + "'" +
                        "   AND  a.line_item_no = '" + ZO.LINE_ITEM_NO + "'" +
                        "   ORDER BY order_no,LINE_ITEM_NO" +
                        "))" +
                        " WHERE part_no = '" + ZO.PART_NO + "'" +
                        " ORDER BY order_no,package_num2";
                    // při úsporném balení vrátí více řádků (kartonů) pro výrobu jedné položky 
                    sPACKAGE_QTY_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "PACKAGE_QTY_MULTI", ref sErr);
                    sNOMINAL_TIME_NA_MULTI = ff.Oracle_GetValue(sConString, sSQL_Multi, "NOMINAL_TIME_NA_MULTI", ref sErr);
                    sDelay_step_multi = ff.Oracle_GetValue(sConString, sSQL_Multi, "delay_step_multi", ref sErr);


                }

                // KARTONY
                // výběr kartonů  ze zakázek z IFS 
                string sSQL_K = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_K = sSQL_K.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_K = sSQL_K.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_K = sSQL_K.Replace("xxTypNalepky", "KTN");
                OracleCommand cmdK = new OracleCommand(sSQL_K, cn);
                OracleDataReader rdK = cmdK.ExecuteReader();
                if (!rdK.HasRows)
                {
                    MessageBox.Show("Nejsou kartony. Jak to ??");
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
                }
                if ((i == 1))
                {
                    sLine = "FILENAME_KTN2 =";
                }
                rdK.Close();
                cmdK.Dispose();


                //  pro Skladacku
                string sSQL_S = Globals.SQL_VyberZeSady.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxTypNalepky", "STN");

                OracleCommand cmdS = new OracleCommand(sSQL_S, cn);
                OracleDataReader rdS = cmdS.ExecuteReader();
                if (!rdS.HasRows)
                {
                    throw new Exception("Nenašel nálepky N-NA-STN !!");
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
                        sAUTO_LABEL = ff.MSSQL_GetOneValue(Globals.ConString_AUSTIN_SLC, "SELECT distinct HodnotaAI FROM NewAI WHERE HodnotaPole = '" + sHodnotaPole + "'");
                        sAUTO_LABEL = "AUTO_LABEL = " + sAUTO_LABEL;

                    }
                    if (rdS["Title"].GetType() != typeof(DBNull))
                    {
                        sTitle = rdS["Title"].ToString();
                    }
                }
                rdS.Close();
                cmdS.Dispose();

                //  pro Praporek
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
                        //" AND a.key_ref LIKE  '%ORDER_NO=" + ZO.ORDER_NO + "%'" +
                        //" AND a.key_ref LIKE  '%LINE_NO=" + ZO.LINE_NO + "%'" +
                        //" AND a.key_ref LIKE  '%REL_NO=" + ZO.RELEASE_NO + "%'" +
                        " AND a.key_value LIKE  '" + ZO.ORDER_NO + "^" + ZO.LINE_NO + "^" + ZO.RELEASE_NO + "^%'" +
                        " AND c.doc_class LIKE '%PRO%'";


                OracleCommand cmdP = new OracleCommand(sSQL_P, cn);
                OracleDataReader rdP = cmdP.ExecuteReader();
                if (!rdP.HasRows) // NOT !!!!!!!!!!!
                {
                    //throw new Exception("Nenašel nálepky PRO !!");
                    MessageBox.Show("Nenašel nálepky PRO !!");
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
                            }
                            if (sPRO.Contains(".BTW") && bPRO_BTW == true && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
                            }
                            if (sPRO.Contains(".ZPL") && bPRO_BTW == false && iRows > 1)
                            {
                                sLine = ("FILENAME_PRO = " + rdP.GetString(0).Trim());
                                sFILENAME_PRO = rdP.GetString(0).Trim();
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
                sSQL_S = Globals.SQL_DOC_CLASS.Replace("xxORDER_NO", ZO.ORDER_NO);
                sSQL_S = sSQL_S.Replace("xxLINE_NO", ZO.LINE_NO);
                sSQL_S = sSQL_S.Replace("xxRELEASE_NO", ZO.RELEASE_NO);
                sSQL_S = sSQL_S.Replace("xxFILE_NAME", _FileNameBTW);
                sErr = ff.Oracle_GetValue3(sConString, sSQL_S, ref sDOC_CLASS, ref sDOC_NO, ref sDOC_REV);


                //CUSTOMER_PO_NO 
                sSQL_1 = "SELECT CUSTOMER_PO_NO FROM IFSAPP.CUSTOMER_ORDER WHERE order_no = '" + ZO.ORDER_NO + "'";
                string sCUSTOMER_PO_NO = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CUSTOMER_PO_NO = " + sCUSTOMER_PO_NO;

                //ORDER_NO
                sLine = "ORDER_NO = " + ZO.ORDER_NO;

                //LINE_NO
                sLine = "LINE_NO = " + ZO.LINE_NO;

                //RELEASE_NO
                sLine = "RELEASE_NO = " + ZO.RELEASE_NO;

                //LINE_ITEM_NO
                sLine = "LINE_ITEM_NO = " + ZO.LINE_ITEM_NO;

                //PACKAGE_NUM1
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM1) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM1 = " + sPACKAGE_NUM1;

                //PACKAGE_NUM2
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2 = " + sPACKAGE_NUM2;

                //CATALOG_NO
                sSQL_1 = "SELECT  catalog_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCatalog_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_NO = " + sCatalog_no;

                //CATALOG_DESC
                sSQL_1 = "SELECT  CATALOG_DESC FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCATALOG_DESC = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CATALOG_DESC = " + sCATALOG_DESC;

                //PACKAGE_QTY
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    // ???
                    //sSQL_1 = "SELECT  IFSAPP.C_Package_Head_Cfp.Get_Cf$_Package_Capacity(PACKING_INSTRUCTION_ID) as package_qty FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                    sSQL_1 = "SELECT  ifsapp.C_Package_Head_Api.Get_Package_Capacity('" + ZO.ORDER_NO + "','" +
                        ZO.LINE_NO + "','" + ZO.RELEASE_NO  + "', -1 , 'CustomerOrderLine') as package_qty FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet; 
                    
                }
                else
                {
                    sSQL_1 = "SELECT  AMOUNT FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }

                string sPACKAGE_QTY = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iPACKAGE_QTY = 0;
                if (Int32.TryParse(sPACKAGE_QTY, out iPACKAGE_QTY)) { } else { MessageBox.Show("Nenašel množství v kartonu !!"); return; };
                sLine = "PACKAGE_QTY = " + sPACKAGE_QTY;

                //SMALL_PACKAGE_QTY 120/4
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    sSQL_1 = "SELECT b.quantity FROM IFSAPP.C_PACKAGE_HEAD_TAB a, IFSAPP.PACKING_INSTRUCTION_NODE_TAB b WHERE  a.packing_instruction_id = b.packing_instruction_id AND b.handling_unit_type_id = 'INNER' AND " + sWhereOrderKomplet;
                }
                else
                {
                    sSQL_1 = "SELECT IFSAPP.C_PACKAGE_HEAD_API.Get_Inner_Package_Count(PACKAGE_TYPE,TEMPLATE_ID) FROM IFSAPP.C_PACKAGE_HEAD WHERE " + sWhereOrderKomplet;
                }

                string sSEQ_NUM = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                int iSEQ_NUM = 0;
                if (Int32.TryParse(sSEQ_NUM, out iSEQ_NUM)) { } else { MessageBox.Show("Nenašel počet vnitřních obalů !!"); return; };
                int iSMALL_PACKAGE_QTY = iPACKAGE_QTY / iSEQ_NUM;
                sLine = "SMALL_PACKAGE_QTY = " + iSMALL_PACKAGE_QTY.ToString();


                // ************************************************************************************************

                //PACKAGE_GROSS_WEIGHT
                // bacha každý karton to může mít jinak - sadování


                /*
                if (Globals.IFS_VERSION == IFS_Versions.IFS_9)
                {
                    // dohledat konkrétní karton
                    sSQL_1 = "SELECT  gross_weight FROM IFSAPP.C_PACKAGE_WEIGHT WHERE " + sWhereOrderLineKomplet;
                    return;
                }
                else
                {
                    sSQL_1 = "SELECT  package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                }

                string sPACKAGE_GROSS_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sPACKAGE_GROSS_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_GROSS_WEIGHT));
                sPACKAGE_GROSS_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_GROSS_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_GROSS_WEIGHT);
                sLine = "PACKAGE_GROSS_WEIGHT = " + sPACKAGE_GROSS_WEIGHT;

                //ROUND_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sROUND_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sROUND_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##}", Convert.ToDouble(sROUND_NET_WEIGHT));
                sROUND_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sROUND_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sROUND_NET_WEIGHT);
                sLine = "ROUND_NET_WEIGHT = " + sROUND_NET_WEIGHT;

                //PACKAGE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sPACKAGE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sPACKAGE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:#####.0}", Convert.ToDouble(sPACKAGE_NET_WEIGHT));
                sPACKAGE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_NET_WEIGHT);
                sLine = "PACKAGE_NET_WEIGHT = " + sPACKAGE_NET_WEIGHT;

                //EXPLOSIVE_NET_WEIGHT
                sSQL_1 = "SELECT  package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sEXPLOSIVE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sEXPLOSIVE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:#####.000}", Convert.ToDouble(sEXPLOSIVE_NET_WEIGHT));
                sEXPLOSIVE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sEXPLOSIVE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sEXPLOSIVE_NET_WEIGHT);
                sLine = "EXPLOSIVE_NET_WEIGHT = " + sEXPLOSIVE_NET_WEIGHT;
                */
                // ************************************************************************************************

                //UN_NO
                sSQL_1 = "select hut.proper_shipping_name_id from IFSAPP.C_PACKAGE_HEAD ph, IFSAPP.HANDLING_UNIT_TYPE hut" +
                " WHERE " + sWhereOrderKomplet +
                " AND hut.handling_unit_type_id = IFSAPP.PACKING_INSTRUCTION_API.Get_Top_Handling_Unit(ph.PACKING_INSTRUCTION_ID)";
                //sSQL_1 = "SELECT  un FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                string sUN = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "UN_NO = " + sUN;

                //ADR_RID_CLASS
                //sSQL_1 = "SELECT class FROM IFSAPP.a0000_zo_baleni_hmotnosti WHERE " + sWhereOrderLineKomplet;
                sSQL_1 = "select psn.hazard_class from PROPER_SHIPPING_NAME psn where psn.PROPER_SHIPPING_NAME_ID = '" + sUN + "'";
                string sADR_RID_CLASS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "ADR_RID_CLASS = " + sADR_RID_CLASS;
                //DETONATOR_CLASS
                string sDETONATOR_CLASS = string.Empty;
                sLine = "DETONATOR_CLASS = "; // u DA není
                                              //sw.WriteLine(sLine);

                //CONSUMPTION_DATE
                sSQL_1 = "SELECT part_no FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sPart_no = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT consumption_month FROM IFSAPP.a0000_polozky_char WHERE part_no = '" + sPart_no + "' AND contract = 'AD'";
                string sConsumption_month = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                sSQL_1 = "SELECT to_char(add_months(wanted_delivery_date," + sConsumption_month + "),'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sCONSUMPTION_DATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CONSUMPTION_DATE = " + sCONSUMPTION_DATE;

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

                //DELAY_STEP = L004,N004
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L004'";
                string sDELAY_STEP = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DELAY_STEP = " + sDELAY_STEP; // u DA není

                //DELAY_STEP_MULTI  
                sLine = "DELAY_STEP_MULTI = "; // u DA není

                // CAVITY_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L003'";
                string sCAVITY_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "CAVITY_MATERIAL  = " + sCAVITY_MATERIAL;

                //SUPPLY_WIRE_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L008'";
                string sSUPPLY_WIRE_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "SUPPLY_WIRE_MATERIAL = " + sSUPPLY_WIRE_MATERIAL;

                //WIRE_ISULATOR_MATERIAL 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'L009'";
                string sWIRE_ISULATOR_MATERIAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_ISULATOR_MATERIAL = " + sWIRE_ISULATOR_MATERIAL;

                //WIRE_LENGTH 
                sSQL_1 = "SELECT b.characteristic_value_desc FROM IFSAPP.INV_PART_DISCRETE_CHAR a,IFSAPP.discrete_charac_value b" +
                    " WHERE a.characteristic_code = b.characteristic_code AND a.attr_value = b.characteristic_value" +
                    " AND  a.part_no = '" + sPart_no + "' AND a.contract = 'AD'" +
                    " AND a.characteristic_code = 'E005'";
                string sWIRE_LENGTH = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "WIRE_LENGTH = " + sWIRE_LENGTH;

                //NOMINAL_TIME_EA 
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sNOMINAL_TIME_EA = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "NOMINAL_TIME_EA = " + sNOMINAL_TIME_EA;

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

                //TESTING_PROTOCOL_NUMBER
                sSQL_1 = "SELECT attr_value FROM IFSAPP.SALES_PART_CHARACTERISTIC WHERE catalog_no = '" + sSales_part + "' AND contract = 'AD'" +
                    " AND  characteristic_code = 'CE011'";
                string sTESTING_PROTOCOL_NUMBER = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TESTING_PROTOCOL_NUMBER = " + sTESTING_PROTOCOL_NUMBER;

                //CUSTOMER_AUTHORIZE_DECREE
                sSQL_1 = "SELECT dt.output_type FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + ZO.ORDER_NO + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sOutput_type = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sSQL_1 = "SELECT dt.note_text FROM OUTPUT_TYPE_DOCUMENT_TAB otd,customer_order_tab co,document_text_tab dt" +
                        " WHERE co.order_no = ' " + ZO.ORDER_NO + "'" +
                        " AND otd.document_code = '99'" +
                        " AND co.note_id = dt.note_id" +
                        " AND otd.output_type = dt.output_type";
                string sNote_text = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                string sCUSTOMER_AUTHORIZE_DECREE = string.Empty;
                switch (sOutput_type)
                {
                    case "LAB01":
                        sLine = "CUSTOMER_AUTHORIZE_DECREE = " + sNote_text;
                        sCUSTOMER_AUTHORIZE_DECREE = sNote_text;
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
                //GENERAL
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GENE'";
                string sGENERAL = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL = " + sGENERAL;
                //GENERAL_1
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'GEN01'";
                string sGENERAL_1 = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "GENERAL_1 = " + sGENERAL_1;
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
                sSQL_1 = "SELECT to_char(wanted_delivery_date,'YYMMDD') FROM IFSAPP.customer_order_line WHERE " + sWhereOrderLineKomplet;
                string sDATE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "DATE = " + sDATE;

                //EXPIRATION
                sLine = "EXPIRATION = " + sCONSUMPTION_DATE;

                //PACKAGE_NUM2BEG
                sSQL_1 = "SELECT  MIN(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2BEG = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2BEG = " + sPACKAGE_NUM2BEG;

                //PACKAGE_NUM2END
                sSQL_1 = "SELECT  MAX(PACKAGE_NUM2) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sPACKAGE_NUM2END = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "PACKAGE_NUM2END = " + sPACKAGE_NUM2END;

                //SEQ_TOTAL - počet skládaček sSMALL_PACKAGE_QTY
                sLine = "SEQ_TOTAL = " + iSEQ_NUM.ToString();

                //BOX_PCS
                sLine = "BOX_PCS = " + iSMALL_PACKAGE_QTY.ToString();

                //TOTAL_PCS
                sSQL_1 = "SELECT SUM(amount) FROM IFSAPP.C_PACKAGE_DETAIL WHERE " + sWhereOrderKomplet;
                string sTOTAL_PCS = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = "TOTAL_PCS = " + sTOTAL_PCS;
                //UFI_CODE
                sSQL_1 = "select ch.CF$_UFI_CODE_A1149 from IFSAPP.A1135_INVPART_CHAR_ALL_CFV ch where ch.CONTRACT = 'AD' and ch.PART_NO = '" + sPart_no + "' AND ch.CHARACTERISTIC_CODE = 'W001'";
                string sUFI_CODE = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);

                //PRODUCER
                sLine = "PRODUCER = " + sPRODUCER;
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

                //DMX_PACKAGE_QTY
                sLine = sDMX_TEXT_PACKAGE_QTY + sGS_DMX_PACKAGE_QTY + sPACKAGE_QTY;
                //

                //DMX_DELAY_LEVEL
                sLine = sDMX_TEXT_DELAY_LEVEL + sGS_DMX_DELAY_LEVEL + sDELAY_LEVEL;

                //DMX_DEALY_STEP ;
                sSQL_1 = "SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR WHERE part_no = '" + sPart_no + "' AND contract = 'AD'" +
                    " AND characteristic_code = 'E016'";
                string sDMX_DEALY_STEP_Value = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                sLine = sDMX_TEXT_DELAY_STEP + sGS_DMX_DELAY_STEP + sDMX_DEALY_STEP_Value;

                //DMX_NOMINAL_TIME_EA
                sLine = sDMX_TEXT_NOMINAL_TIME_EA + sGS_DMX_NOMINAL_TIME_EA + sNOMINAL_TIME_EA;

                //DMX_NOMINAL_TIME_NA
                sLine = sDMX_TEXT_NOMINAL_TIME_NA + sGS_DMX_NOMINAL_TIME_NA + sNOMINAL_TIME_NA;

                //DMX_ST_LENGTH
                sLine = sDMX_TEXT_ST_LENGTH + sGS_DMX_ST_LENGTH + sST_LENGTH;

                //DMX_DATE
                sLine = sDMX_TEXT_DATE + sGS_DMX_DATE + sDATE;

                //DMX_PACKAGE_NUM2
                sLine = sDMX_TEXT_PACKAGE_NUM2 + sGS_DMX_PACKAGE_NUM2 + sPACKAGE_NUM2;

                //DMX_SEQ_NUM
                string sDMX_SEQ_NUM = iSEQ_NUM.ToString();
                sLine = sDMX_TEXT_SEQ_NUM + sGS_DMX_SEQ_NUM + sDMX_SEQ_NUM;

                //DMX_WIRE_LENGTH
                string sDMX_WIRE_LENGTH = sWIRE_LENGTH;
                sLine = sDMX_TEXT_WIRE_LENGTH + sGS_DMX_WIRE_LENGTH + sDMX_WIRE_LENGTH;

                //AUTO_LABEL - nové aplikační identifikátory (je to v SL), podle nich si TRC vybírá praporek
                //sw.WriteLine(sAUTO_LABEL);
                //
                //if (checkBoxDETONATOR_APLIC.Checked) // doplněno 22.9.2014 Senkerik
                //{
                //    sLine = "DETONATOR_APPLIC = NO";
                //    sw.WriteLine(sLine);
                //}
                // datatable
                int iBeg = 0;
                int iEnd = 0;
                int iPocetKartonu = 0;
                int iPocetSkladacek = 0;
                Int32.TryParse(sPACKAGE_NUM2BEG, out iBeg);
                Int32.TryParse(sPACKAGE_NUM2END, out iEnd);
                int iOpakovani = (iUroven == 1 ? 1 : int.Parse(sSEQ_NUM));
                for (int ii = iBeg; ii <= iEnd; ii++) // kartony
                {
                    for (int iU = 1; iU <= iOpakovani; iU++) // úroveň
                    {

                        DataRow dr = _dtNalepky.NewRow();

                        sPackageBarcodeId = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString();
                        //if (sPackageBarcodeId == "V92455-2-1-51")
                        //{
                        //    Console.WriteLine("");
                        //}
                        // dohledat hmotnosti pro karton
                        sSQL_1 = "SELECT package_weight_brut FROM IFSAPP.a0000_zo_baleni_hmotnosti cpw WHERE " + sWhereOrderKomplet +
                            " AND PACKAGE_NUM2 = " + ii.ToString();
                        string sPACKAGE_GROSS_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                        sPACKAGE_GROSS_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##.0}", Convert.ToDouble(sPACKAGE_GROSS_WEIGHT));
                        sPACKAGE_GROSS_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_GROSS_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_GROSS_WEIGHT);

                        //ROUND_PACKAGE_NET_WEIGHT, sPACKAGE_NET_WEIGHT
                        sSQL_1 = "SELECT package_weight_net FROM IFSAPP.a0000_zo_baleni_hmotnosti cpw WHERE " + sWhereOrderKomplet +
                                 " AND PACKAGE_NUM2 = " + ii.ToString();
                        string sROUND_PACKAGE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                        string sPACKAGE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:#####.0}", Convert.ToDouble(sROUND_PACKAGE_NET_WEIGHT));
                        sPACKAGE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sPACKAGE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sPACKAGE_NET_WEIGHT);

                        sROUND_PACKAGE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:##}", Convert.ToDouble(sROUND_PACKAGE_NET_WEIGHT));
                        sROUND_PACKAGE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sROUND_PACKAGE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sROUND_PACKAGE_NET_WEIGHT);



                        //EXPLOSIVE_NET_WEIGHT
                        sSQL_1 = "SELECT package_weight_explosive_net FROM IFSAPP.a0000_zo_baleni_hmotnosti cpw WHERE " + sWhereOrderKomplet +
                                 " AND PACKAGE_NUM2 = " + ii.ToString();
                        string sEXPLOSIVE_NET_WEIGHT = ff.Oracle_GetValue(sConString, sSQL_1, ref sErr);
                        sEXPLOSIVE_NET_WEIGHT = String.Format(CultureInfo.CurrentCulture, "{0:0.0000}", Convert.ToDouble(sEXPLOSIVE_NET_WEIGHT));
                        sEXPLOSIVE_NET_WEIGHT = (Globals.DECIMAL_SEPARATOR == DecimalSeparator.Dot ? sEXPLOSIVE_NET_WEIGHT.Replace(DecimalSeparator.Comma, DecimalSeparator.Dot) : sEXPLOSIVE_NET_WEIGHT);
                        //
                        dr["PACKAGE_ID"] = sPackageBarcodeId;
                        dr["SUBPACKAGE_ID"] = sPackageBarcodeId + "-" + iU.ToString();


                        dr[BtwHdr.c01_DOC_CLASS] = _DOC_CLASS;
                        dr[BtwHdr.c02_TITLE] = sTitle;
                        dr[BtwHdr.c03_DOC_NO] = sDOC_NO;
                        dr[BtwHdr.c04_DOC_REV] = sDOC_REV;
                        dr[BtwHdr.c05_FILE_NAME] = _FileNameBTW;
                        dr[BtwHdr.c06_CUSTOMER_PO_NO] = sCUSTOMER_PO_NO;
                        dr[BtwHdr.c07_ORDER_NO] = ZO.ORDER_NO;
                        dr[BtwHdr.c08_LINE_NO] = ZO.LINE_NO;
                        dr[BtwHdr.c09_RELEASE_NO] = ZO.RELEASE_NO;
                        dr[BtwHdr.c10_LINE_ITEM_NO] = ZO.LINE_ITEM_NO;
                        dr[BtwHdr.c11_PACKAGE_NUM1] = sPACKAGE_NUM1;
                        dr[BtwHdr.c12_SEQ_NUM] = iU.ToString();
                        dr[BtwHdr.c13_PACKAGE_NUM2] = ii.ToString(); //sPACKAGE_NUM2;
                        dr[BtwHdr.c14_CATALOG_NO] = sCatalog_no;
                        dr[BtwHdr.c15_CATALOG_DESC] = sCATALOG_DESC;
                        dr[BtwHdr.c16_VENDOR_COMPANY] = sVENDOR_COMPANY;
                        dr[BtwHdr.c17_VENDOR_NAME] = sVENDOR_NAME;
                        dr[BtwHdr.c18_VENDOR_ADDRESS_ID] = sVENDOR_ADDRESS_ID;
                        dr[BtwHdr.c19_VENDOR_STREET] = sVENDOR_ADDRESS;
                        dr[BtwHdr.c20_VENDOR_CITY] = sVENDOR_CITY;
                        dr[BtwHdr.c21_VENDOR_ZIP_CODE] = sVENDOR_ZIP_CODE;
                        dr[BtwHdr.c22_VENDOR_COUNTRY] = sVENDOR_COUNTRY;
                        dr[BtwHdr.c23_RECEIVER_COMPANY_NAME2] = sReceiverCompanyName;
                        dr[BtwHdr.c24_RECEIVER_ADDRESS1] = sReceiverAddress1;
                        dr[BtwHdr.c25_RECEIVER_ADDRESS2] = sReceiverAddress2;
                        dr[BtwHdr.c26_RECEIVER_ZIP_CODE] = sReceiverZIP_CODE;
                        dr[BtwHdr.c27_RECEIVER_CITY] = sReceiverCity;
                        dr[BtwHdr.c28_RECEIVER_COUNTRY] = sReceiverCountry;
                        dr[BtwHdr.c29_RECEIVER_NAME] = sReceiverContactName;
                        dr[BtwHdr.c30_RECEIVER_PHONE] = sReceiverContactPhone;
                        dr[BtwHdr.c31_RECEIVER_FAX] = sReceiverContactFax;
                        dr[BtwHdr.c32_RECEIVER_MOBILE] = sReceiverContactMobile;
                        dr[BtwHdr.c33_RECEIVER_EMAIL] = sReceiverContactEmail;
                        dr[BtwHdr.c34_UNIQUE_PACKAGE_ID] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString();
                        dr[BtwHdr.c35_UNIQUE_SUBPACKAGE_ID] = ZO.ORDER_NO + "-" + ZO.LINE_NO + "-" + ZO.RELEASE_NO + "-" + ii.ToString() + "-" + iU.ToString();
                        dr[BtwHdr.c36_PACKAGE_QTY] = sPACKAGE_QTY;
                        dr[BtwHdr.c37_SMALL_PACKAGE_QTY] = iSMALL_PACKAGE_QTY.ToString();
                        dr[BtwHdr.c38_PACKAGE_GROSS_WEIGHT] = sPACKAGE_GROSS_WEIGHT; //string.Empty; //sPACKAGE_GROSS_WEIGHT;
                        dr[BtwHdr.c39_ROUND_NET_WEIGHT] = sROUND_PACKAGE_NET_WEIGHT; //string.Empty; //sPACKAGE_GROSS_WEIGHT;
                        dr[BtwHdr.c40_PACKAGE_NET_WEIGHT] = sPACKAGE_NET_WEIGHT;
                        dr[BtwHdr.c41_EXPLOSIVE_NET_WEIGHT] = sEXPLOSIVE_NET_WEIGHT;
                        dr[BtwHdr.c42_UN_NO] = sUN;
                        dr[BtwHdr.c43_ADR_RID_CLASS] = sADR_RID_CLASS;
                        dr[BtwHdr.c44_DETONATOR_CLASS] = sDETONATOR_CLASS;
                        dr[BtwHdr.c45_DELAY_LEVEL] = sDELAY_LEVEL;
                        dr[BtwHdr.c46_CAVITY_MATERIAL] = sCAVITY_MATERIAL;
                        dr[BtwHdr.c47_IGNITION_CUE_MJ] = string.Empty;
                        dr[BtwHdr.c48_IGNITION_CUE_J] = string.Empty;
                        dr[BtwHdr.c49_SAFETY_STREAM] = string.Empty;
                        dr[BtwHdr.c50_SAFETY_CUE_MJ] = string.Empty;
                        dr[BtwHdr.c51_SAFETY_CUE_J] = string.Empty;
                        dr[BtwHdr.c52_SUPPLY_WIRE_MATERIAL] = sSUPPLY_WIRE_MATERIAL;
                        dr[BtwHdr.c53_WIRE_ISULATOR_MATERIAL] = sWIRE_ISULATOR_MATERIAL;
                        dr[BtwHdr.c54_WIRE_LENGTH] = sWIRE_LENGTH;
                        dr[BtwHdr.c55_WIRE_LENGTH_TRACE] = string.Empty;
                        dr[BtwHdr.c56_CONSUMPTION_DATE] = sCONSUMPTION_DATE;
                        dr[BtwHdr.c57_NOMINAL_TIME_EA] = sNOMINAL_TIME_EA;
                        dr[BtwHdr.c58_NOMINAL_TIME_NA] = sNOMINAL_TIME_NA;
                        dr[BtwHdr.c59_ST_LENGTH] = sST_LENGTH;
                        dr[BtwHdr.c60_DELAY_STEP] = sDELAY_STEP;
                        dr[BtwHdr.c61_BATCH_NO] = sBatchNo;
                        dr[BtwHdr.c62_TESTING_STATION_NUMBER] = sTESTING_STATION_NUMBER;
                        dr[BtwHdr.c63_TESTING_PROTOCOL_NUMBER] = sTESTING_PROTOCOL_NUMBER;
                        dr[BtwHdr.c64_CUSTOMER_AUTHORIZE_DECREE] = sCUSTOMER_AUTHORIZE_DECREE;
                        dr[BtwHdr.c65_ANOTHER_NOTICE] = sNote_text;
                        dr[BtwHdr.c66_NOTE] = sNote_text;
                        dr[BtwHdr.c67_OBJVERSION] = string.Empty; // oracle sloupec
                        dr[BtwHdr.c68_OBJID] = string.Empty; // oracle sloupec 
                        dr[BtwHdr.c69_DATE] = sDATE;
                        dr[BtwHdr.c70_EXPIRATION] = sCONSUMPTION_DATE;
                        dr[BtwHdr.c71_TOTAL_RESIST] = string.Empty;
                        dr[BtwHdr.c72_BRIDGE_RESIST] = string.Empty;
                        dr[BtwHdr.c73_DATE_CODE] = string.Empty;
                        dr[BtwHdr.c74_LOT_NUMBER] = string.Empty;
                        dr[BtwHdr.c75_STATION_NUM] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c76_WORKERS] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c77_WORKER2] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c78_DUPLICATE] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c79_PIECE_NUM] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c80_DELAY_FRACTION] = string.Empty; // není v c_labels_ad;
                        dr[BtwHdr.c81_GENERAL] = sGENERAL;
                        dr[BtwHdr.c82_GENERAL_1] = sGENERAL_1;
                        dr[BtwHdr.c83_BOX_PCS] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c84_MANUAL] = string.Empty; // není v c_labels_ad
                        dr[BtwHdr.c85_PRODUCER] = sPRODUCER;
                        dr[BtwHdr.c86_DMX_PACKAGE_QTY] = sPACKAGE_QTY;
                        dr[BtwHdr.c87_DMX_DELAY_LEVEL] = sDELAY_LEVEL;
                        dr[BtwHdr.c88_DMX_DELAY_STEP] = sDELAY_STEP;
                        dr[BtwHdr.c89_DMX_NOMINAL_TIME_EA] = sNOMINAL_TIME_EA;
                        dr[BtwHdr.c90_DMX_NOMINAL_TIME_NA] = sNOMINAL_TIME_NA;
                        dr[BtwHdr.c91_DMX_WIRE_LENGTH] = sWIRE_LENGTH;
                        dr[BtwHdr.c92_DMX_ST_LENGTH] = sST_LENGTH;
                        dr[BtwHdr.c93_DMX_DATE] = sDATE;
                        dr[BtwHdr.c94_DMX_PACKAGE_NUM2] = sPACKAGE_NUM2;
                        dr[BtwHdr.c95_DMX_SEQ_NUM] = sSEQ_NUM;
                        dr[BtwHdr.c96_DMX_PRODUCER] = sPRODUCER;
                        dr[BtwHdr.c97_AUTO_LABEL] = sAUTO_LABEL;
                        dr[BtwHdr.c98_TOTAL_PACKAGE_QTY] = sTOTAL_PCS;
                        dr[BtwHdr.c99_UFI_CODE] = sUFI_CODE;


                        _dtNalepky.Rows.Add(dr);
                        iPocetSkladacek++;
                    }
                    iPocetKartonu++;
                }
                DataRow dr1 = _dtNalepky.NewRow();
                dr1["PACKAGE_ID"] = "Počet:" + iPocetKartonu.ToString();
                dr1["SUBPACKAGE_ID"] = "Počet:" + iPocetSkladacek.ToString();
                _dtNalepky.Rows.Add(dr1);

                dataGridViewLabels.DataSource = _dtNalepky;
                dataGridViewLabels.Refresh();
                for (int jj = 1; jj <= iSEQ_NUM; jj++)
                {
                }
                //

                //Napln_dataGridViewSoubory();
                cn.Close();
                //sw.Dispose();
                //sw.Close();
                //MessageBox.Show("Soubor je vytvořen.");
                //KopirujNalepky();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+ " " + sPackageBarcodeId);
            }
            finally
            {
            }
            this.Cursor = Cursors.Default;
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
        private string strSplit(string pText, string pField,string pDelimiter)
        {
            string sPom = string.Empty;
            string[] subs = pText.Split(';');

            foreach (var sub in subs)
            {
                if (sub.Contains(pField))
                {
                    string[] list2 = sub.Replace(pField,"").Split('.');

                    foreach (var aa in list2)
                    {
                        if (aa.Trim() != string.Empty)
                        {
                            sPom = sPom + aa + ".";
                        }
                    }
                }
            }

            return sPom;

        }
    }
}
