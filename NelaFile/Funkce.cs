using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
//using Tamir.SharpSsh;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Data.OracleClient;
using System.Data.SqlClient;

//using Ionic.Zip;

namespace NelaFile
{

    public class Funkce
    {
        public string ConString_MSSQL_Austin_Trc { get; set; }

        private const string KEY_QUOTE = "\"";
        //	private string name;
        //	public Funkce()
        //	{
        //		//
        //		// TODO: Add constructor logic here
        //		//
        //		name = "Funkce";
        //	}
        public string sefStringReplace(string pText, string pPuvodni, string pNovy)
        {
            //pPuvodni = AddKEY_QUOTEs(pPuvodni);
            //pNovy = AddKEY_QUOTEs(pNovy);
            string sPom = pText.Replace(pPuvodni, pNovy);
            return sPom;
        }

        public string[] sefStringDejPole(string pText, char pOddelovac)
        {
            char[] delimiters = new char[] { pOddelovac, pOddelovac };
            string[] sPom;
            sPom = pText.Split(delimiters);
            return sPom;
        }

        private string AddKEY_QUOTEs(string pValue)
        {
            return KEY_QUOTE + pValue.Replace(KEY_QUOTE, KEY_QUOTE + KEY_QUOTE) + KEY_QUOTE;
        }
        public string DejHodnotu_INI(string pFile, string pSekce)
        {
            //StreamReader myStreamReader = null;
            //TextReader myStreamReader = null;
            string myInputstring;
            int rowCount = 0;
            int iDatabase = 0;
            const string sREM = "REM";
            string sPom;
            string sCesta = "";

            StreamReader myStreamReader = new StreamReader(
                (System.IO.Stream)File.OpenRead(@pFile), System.Text.Encoding.Default);

            try
            {
                //TextReader myStreamReader;

                //myStreamReader = File.OpenText(@pFile);

                //myStreamReader = new StreamReader(pFile,System.Text.Encoding.Unicode); // (System.IO.TextReader) File.Open(pFile,System.IO.FileMode.Open));				
                //myStreamReader = new StreamReader((System.IO.Stream) File.OpenRead( pFile,System.Text.Encoding.ASCII));
                //myStreamReader = new StreamReader(( File.OpenText(pFile),System.Text.Encoding.Unicode);

                myInputstring = myStreamReader.ReadLine();
                while (myInputstring != null)
                {
                    if (iDatabase == 1)
                    {
                        sPom = myInputstring.ToUpper().Substring(0, 3).ToString();
                        if (sPom != sREM)
                        {
                            sCesta = myInputstring;
                            iDatabase = 0;
                        }
                    }
                    if (myInputstring.ToUpper() == pSekce.ToUpper())
                    {
                        iDatabase = 1;
                    }
                    rowCount++;
                    // Read the next line.
                    myInputstring = myStreamReader.ReadLine();
                }

            }
            catch (Exception exc)
            {
                // Show the error to the user.
                MessageBox.Show("File could not be opened or read." + Environment.NewLine + "Please verify that the filename is correct, and that you have read permissions for the desired directory." + Environment.NewLine + Environment.NewLine + "Exception: " + exc.Message);
            }
            finally
            {
                // Close the object if it has been created.
                if (myStreamReader != null)
                {
                    myStreamReader.Close();
                }
            }
            return sCesta;
        }
        public void INI_UlozHodnoty(
            string pFile,
            string pDatabaseS, string pDatabaseH,
            string pNazevDbS, string pNazevDbH,
            string pFolderForExportS, string pFolderForExportH,
            string pExePdfS, string pExePdfH
            )
        {
            using (StreamWriter sw = new StreamWriter(pFile))
            {
                // Add some text to the file.
                sw.WriteLine(pDatabaseS);
                sw.WriteLine(pDatabaseH);
                //
                sw.WriteLine(pNazevDbS);
                sw.WriteLine(pNazevDbH);
                //
                sw.WriteLine(pFolderForExportS);
                sw.WriteLine(pFolderForExportH);
                //
                sw.WriteLine(pExePdfS);
                sw.WriteLine(pExePdfH);
                //
                sw.Close();
            }
        }
        public string FileTxt_NajdivCSV(string pFile,
                string pDelimiter,
              string pHledej, int iIndexHledej, int iIndexVrat, int iIndexDalsi)
        {

            string myInputstring;
            int rowCount = 0;
            int i = 0;
            //string sCesta = "";

            StreamReader myStreamReader = new StreamReader(
                (System.IO.Stream)File.OpenRead(@pFile), System.Text.Encoding.Default);

            try
            {
                myInputstring = myStreamReader.ReadLine();
                while (myInputstring != null)
                {
                    string[] words = Regex.Split(myInputstring, pDelimiter);

                    foreach (string word in words)
                    {
                        if (i == iIndexHledej)
                        {
                            if (word == pHledej)
                            {
                                myStreamReader.Close();
                                if (iIndexDalsi > -1)
                                {
                                    return words[iIndexVrat].ToString() + " " + words[iIndexDalsi].ToString();
                                }
                                else
                                    return words[iIndexVrat].ToString();
                            }
                        }
                        i++;
                    }
                    i = 0;
                    rowCount++;
                    // Read the next line.
                    myInputstring = myStreamReader.ReadLine();
                }

            }
            catch (Exception exc)
            {
                // Show the error to the user.
                MessageBox.Show("File could not be opened or read." + Environment.NewLine + "Please verify that the filename is correct, and that you have read permissions for the desired directory." + Environment.NewLine + Environment.NewLine + "Exception: " + exc.Message);
            }
            finally
            {
                // Close the object if it has been created.
                if (myStreamReader != null)
                {
                    myStreamReader.Close();
                }
            }
            return string.Empty;
        }
        public bool PristupPovolen(string pUser)
        {
            if (pUser.ToLower() == "senkerik" || pUser.ToLower() == "fojtik" || pUser.ToLower() == "hadas")
                return true;
            else
                return false;

        }
        public bool PristupPovolen_DeleteDocumentCode(string pUser)
        {
            if (pUser.ToLower() == "zezulkova" || pUser.ToLower() == "dancak" || pUser.ToLower() == "vaculikova")
                return true;
            else
                return false;

        }

        public string File_DejName(string pFile)
        {

            string sZakazka = XML_GetElements(pFile, "/dsMessage/PurchaseOrder");
            string sFile = "CZ010_" + sZakazka + ".xml";
            return sFile;
        }

        public string XML_GetElements(string pXMLFile, string pNode)
        {

            //"/catalog/cd/price"

            string fileName = pXMLFile;
            string sValue = string.Empty;
            XPathDocument doc = new XPathDocument(pXMLFile);
            XPathNavigator nav = doc.CreateNavigator();

            // Compile a standard XPath expression
            XPathExpression expr;
            expr = nav.Compile(pNode);
            XPathNodeIterator iterator = nav.Select(expr);

            // Iterate on the node set
            //listBox1.Items.Clear();
            try
            {
                while (iterator.MoveNext())
                {
                    XPathNavigator nav2 = iterator.Current.Clone();
                    sValue = nav2.Value;

                    //listBox1.Items.Add("price: " + nav2.Value);

                }
                return sValue;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }

        }
        public bool Zip(string pFile,ref string pMsgError)
        {
            try
            {
            //    string args = " -D " + pFile + ".zip " + pFile;

            //    string sZipExe = AppDomain.CurrentDomain.BaseDirectory + "zip.exe";

            //    try
            //    {
            //        using (ZipFile zip = new ZipFile())
            //        {
            //            ZipEntry e = zip.AddFile(pFile,@"\");
            //            e.Comment = "Added by Cheeso's CreateZip utility.";
            //            zip.Comment = String.Format("This zip archive was created by the CreateZip example application on machine '{0}'",
            //               System.Net.Dns.GetHostName());
            //            zip.Save(pFile + ".zip");
            //        }

            //    }
            //    catch (System.Exception ex1)
            //    {
            //        pMsgError = ex1.Message.ToString();
            //        return false;
            //    }


                return true;
            }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
            catch (Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
            {
            //    pMsgError = ex.Message.ToString();
               return false;
            }
        }
        public Boolean sftp_copy(string pSource, string pDestination)
        {
            try
            {

                //Tamir.SharpSsh.Sftp ss = new Sftp("212.24.149.193", "austin", "r6Gwkt73dE");
                //ss.Connect();
                //ss.Put(pSource, pDestination); //"/DE-KaS/aa.xml")
                //ss.Close();
                return true;
            }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
            catch (Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
            {
                //MessageBox.Show(ex.Message);
                return false;
            }

        }
        public DialogResult InputBox(string title, string promptText, ref string pValue)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = pValue;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            pValue = textBox.Text;
            return dialogResult;
        }

        #region Oracle
        public string Oracle_GetValue(string pConString, string pSQL, ref string pErr)
        {
            string sText = string.Empty;
            pErr = string.Empty;
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            switch (reader[0].GetType().ToString())
                            {
                                case "System.Decimal":
                                    sText = Convert.ToString(reader.GetDecimal(0)); break;
                                case "System.DateTime":
                                    sText = Convert.ToString(reader.GetDateTime(0).ToShortDateString()); break;
                                default:
                                    sText = reader.GetString(0).ToString(); break;
                            }
                        }
                        break;
                    }

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                pErr = ex.Message.ToString();
                return "";
            }
            finally
            {
            }
            return sText;
        }
        public string Oracle_GetValue(string pConString, string pSQL,string pSloupec, ref string pErr)
        {
            string sText = string.Empty;
            pErr = string.Empty;
#pragma warning disable CS0219 // The variable 'i' is assigned but its value is never used
            int i = 0;
#pragma warning restore CS0219 // The variable 'i' is assigned but its value is never used
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    //DataTable dt = new DataTable();
                    //dt.Load(reader);
                    //dt.Rows.Count
                    while (reader.Read())
                    {
                        if (reader[pSloupec].GetType() != typeof(DBNull))
                        {
                            sText = reader[pSloupec].ToString();
                        }
                        //break;
                    }
                    //dt.Dispose();
                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                pErr = ex.Message.ToString();
                return "";
            }
            finally
            {
            }
            return sText;
        }
        public string Oracle_Execute(string pConString, string pSQL)
        {
            string sErr = string.Empty;
            try
            {
                string sSQL = pSQL;

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                //
                cmd.ExecuteNonQuery();
                //
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                sErr = ex.Message.ToString();
                return sErr;
            }
            finally
            {
            }
            return string.Empty;
        }
        public string Oracle_GetValue2(string pConString, string pSQL, ref string pHodnota1, ref string pHodnota2)
        {
            string sText = string.Empty;
            string sErr = string.Empty;

            //pHodnota1 = posledn sklad prijmu
            //pHodnota2 = transaction_id nebo transaction_id_ad
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            pHodnota1 = reader.GetString(0);
                        }
                        if (reader[1].GetType() != typeof(DBNull))
                        {
                            switch (reader[1].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota2 = Convert.ToString(reader.GetDecimal(1)); break;
                                default:
                                    pHodnota2 = reader.GetString(1).ToString(); break;
                            }
                        }
                        if (reader[2].GetType() != typeof(DBNull))
                        {
                            switch (reader[2].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota2 = Convert.ToString(reader.GetDecimal(2)); break;
                                default:
                                    pHodnota2 = reader.GetString(2).ToString(); break;
                            }
                        }
                        break;
                    }

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                sErr = ex.Message.ToString();
                return sErr;
            }
            finally
            {
            }
            return string.Empty;
        }
        public string Oracle_GetValue3(string pConString, string pSQL, ref string pHodnota1, ref string pHodnota2, ref string pHodnota3)
        {
            string sText = string.Empty;
            string sErr = string.Empty;

            try
            {
                string sSQL = pSQL;

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            switch (reader[0].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota1 = Convert.ToString(reader.GetDecimal(0)); break;
                                default:
                                    pHodnota1 = reader.GetString(0).ToString(); break;
                            }
                        }
                        if (reader[1].GetType() != typeof(DBNull))
                        {
                            switch (reader[1].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota2 = Convert.ToString(reader.GetDecimal(1)); break;
                                default:
                                    pHodnota2 = reader.GetString(1).ToString(); break;
                            }
                        }
                        if (reader[2].GetType() != typeof(DBNull))
                        {
                            switch (reader[2].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota3 = Convert.ToString(reader.GetDecimal(2)); break;
                                default:
                                    pHodnota3 = reader.GetString(2).ToString(); break;
                            }
                        }
                        break;
                    }

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                sErr = ex.Message.ToString();
                return sErr;
            }
            finally
            {
            }
            return string.Empty;
        }
        public bool Oracle_Exists(string pConString, string pSQL, ref string pErr)
        {
            bool bExistuje = false;
            pErr = string.Empty;
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    bExistuje = true;
                    //while (reader.Read())
                    //{
                    //    if (reader[0].GetType() != typeof(DBNull))
                    //    {
                    //        sText = reader.GetString(0);
                    //    }
                    //    break;
                    //}

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                pErr = ex.Message.ToString();
                return bExistuje;
            }
            finally
            {
            }
            return bExistuje;
        }
        #endregion
        #region MSSQL
        public string MSSQL_Execute(string pSQL)
        {
            //string cnsVSQ01 = "Persist Security Info=True;User ID=trc;Password=123456;Initial Catalog=AUSTIN_TRC;Data Source=vsq01";
            try
            {
                SqlConnection cnSQL01 = new SqlConnection(ConString_MSSQL_Austin_Trc); ;
                cnSQL01.Open();

                //'2014-11-24 9:02:57'

                SqlCommand cmd = cnSQL01.CreateCommand();
                cmd.CommandText = pSQL;
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                cnSQL01.Close();
                return "";
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }

        }
        public string MSSQL_GetOneValue(string pConnectString, string pSQL)
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

        #endregion

    }
    class csVerze
    {
        //This string represents the person's Name.
        public string Name { get; set; }

        //This string represents the person's Email.
        public string Email { get; set; }

        public bool ExistujeNovaVerze(string pFileOld, string pFileNew, ref string pVerzeOld, ref string pVerzeNew)
        {
            //"\\\\Vfs01\\PRENOSY\\XML\\plxml.exe"

            string sVerzeOld = "";
            string sVerzeNew = "";

#pragma warning disable CS0219 // The variable 'iVersionServer' is assigned but its value is never used
            Int64 iVersionServer = 0;
#pragma warning restore CS0219 // The variable 'iVersionServer' is assigned but its value is never used
#pragma warning disable CS0219 // The variable 'iVersion' is assigned but its value is never used
            Int64 iVersion = 0;
#pragma warning restore CS0219 // The variable 'iVersion' is assigned but its value is never used


            Int64 iVServer_Major = 0;
            Int64 iV_Major = 0;


            Int64 iVServer_Minor = 0;
            Int64 iV_Minor = 0;

            Int64 iVServer_bp = 0;
            Int64 iV_bp = 0;

            Int64 iVServer_fp = 0;
            Int64 iV_fp = 0;

            Int64 iVs = 0;
            Int64 iV = 0;
            Int64 iPom = 1000000;
            bool bNovaVerzeExistuje = false;
            FileVersionInfo versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(pFileNew);
            FileVersionInfo myBuildInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(pFileOld);
            sVerzeNew = versionInfo.FileVersion.ToString();
            sVerzeOld = myBuildInfo.FileVersion.ToString();
            pVerzeOld = sVerzeOld;
            pVerzeNew = sVerzeNew;

            //MessageBox.Show(sVerzeOld + "   " + sVerzeNew);
            //gfdgdg
            //gdfgd


            iVServer_Major = versionInfo.FileMajorPart;
            iV_Major = myBuildInfo.FileMajorPart;

            iVServer_Minor = versionInfo.FileMinorPart;
            iV_Minor = myBuildInfo.FileMinorPart;

            iVServer_bp = versionInfo.FileBuildPart;
            iV_bp = myBuildInfo.FileBuildPart;

            iVServer_fp = versionInfo.FilePrivatePart;
            iV_fp = myBuildInfo.FilePrivatePart;

            iVs = iVServer_Major * (iVServer_Minor == 0 ? 1 : iVServer_Minor) * (iVServer_bp * iPom) + iVServer_fp;
            iV = iV_Major * (iV_Minor == 0 ? 1 : iV_Minor) * (iV_bp * iPom) + iV_fp;

            //
            if (iVs > iV)
            {
                bNovaVerzeExistuje = true;
            }

            //if (iVServer_Major > iV_Major)
            //{
            //    MessageBox.Show(sVerzeOld + " " + sVerzeNew);
            //    bNovaVerzeExistuje = true;
            //}
            //else
            //{
            //    if (iVServer_Minor > iV_Minor)
            //    {bNovaVerzeExistuje = true;}
            //    else
            //    {
            //        if (iVServer_bp > iV_bp)
            //        { bNovaVerzeExistuje = true; }
            //        else 
            //        {
            //            if (iVServer_fp > iV_fp && iVServer_bp == iV_bp)
            //            {
            //                MessageBox.Show(sVerzeOld + " " + sVerzeNew);
            //                bNovaVerzeExistuje = true; 
            //            }
            //        }
            //    }

            //}
            return bNovaVerzeExistuje;
        }

        public bool FilesCompareByDate(string pFileOld, string pFileNew)
        {
            //@"d:\MyTextFile.txt"
            DateTime dFileTimeOld = File.GetLastWriteTime(pFileOld);
            DateTime dFileTimeNew = File.GetLastWriteTime(pFileNew);

            if (dFileTimeNew > dFileTimeOld)
            {
                // nový je novìjší
                return true;
            }

            return false;
        }
        public void cmdCopyFile(string sFileOld, string sFileNew)
        {

            Process process = new System.Diagnostics.Process();
            ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            startInfo.UseShellExecute = true;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C xcopy /Y " + sFileNew + " " + sFileOld;

            //MessageBox.Show(startInfo.Arguments.ToString());

            process.StartInfo = startInfo;
            process.Start();
            Application.Exit();
        }
        public void SendEmail(string pSubject, string pBody, string sSMTP_Server)
        {
            try
            {
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                message.To.Add("miroslav.senkerik@austin.cz");
                message.Subject = System.Windows.Forms.SystemInformation.UserName.ToString() + " - " + pSubject; //Application.ProductName.ToString();
                message.From = new System.Net.Mail.MailAddress("From@" + Application.ProductName.ToString() + ".cz");
                message.Body = pBody;
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(sSMTP_Server);
                smtp.EnableSsl = false;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            { }
        }
        public static string StringFormat_NoTruncate(string pString, int pDecimalPlaces)
        {
            char dec = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
            return pString.Substring(0, pString.LastIndexOf(dec) + pDecimalPlaces + 1);
        }

        #region Oracle
        public string Oracle_GetValue(string pConString, string pSQL, ref string pErr)
        {
            string sText = string.Empty;
            pErr = string.Empty;
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            sText = reader.GetString(0);
                        }
                        break;
                    }

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                pErr = ex.Message.ToString();
                return "";
            }
            finally
            {
            }
            return sText;
        }
        public string Oracle_Execute(string pConString, string pSQL)
        {
            string sErr = string.Empty;
            try
            {
                string sSQL = pSQL;

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                //
                cmd.ExecuteNonQuery();
                //
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                sErr = ex.Message.ToString();
                return sErr;
            }
            finally
            {
            }
            return string.Empty;
        }
        public string Oracle_GetValue2(string pConString, string pSQL, ref string pHodnota1, ref string pHodnota2)
        {
            string sText = string.Empty;
            string sErr = string.Empty;

            //pHodnota1 = posledn sklad prijmu
            //pHodnota2 = transaction_id nebo transaction_id_ad
            try
            {
                //Zpùsob komunikace s názvem … sftpWN
                string sSQL = pSQL;
                //"select value from IFSAPP.CUSTOMER_INFO_COMM_METHOD a WHERE a.method_id_DB = 'INTERCOM' AND a.customer_id ='" + _ZakaznikCislo + "'";

                OracleConnection conn = new OracleConnection(pConString);
                conn.Open();
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = sSQL;
                OracleDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if (reader[0].GetType() != typeof(DBNull))
                        {
                            pHodnota1 = reader.GetString(0);
                        }
                        if (reader[1].GetType() != typeof(DBNull))
                        {
                            switch (reader[1].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota2 = Convert.ToString(reader.GetDecimal(1)); break;
                                default:
                                    pHodnota2 = reader.GetString(1).ToString(); break;
                            }
                        }
                        if (reader[2].GetType() != typeof(DBNull))
                        {
                            switch (reader[2].GetType().ToString())
                            {
                                case "System.Decimal":
                                    pHodnota2 = Convert.ToString(reader.GetDecimal(2)); break;
                                default:
                                    pHodnota2 = reader.GetString(2).ToString(); break;
                            }
                        }
                        break;
                    }

                }
                else
                {
                    //MessageBox.Show("Zákazník :" + _ZakaznikCislo + " nemá nastavenou komunikaèní metodu Interkom !!"); 
                }
                reader.Close();
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
                //
            }
            catch (Exception ex)
            {
                sErr = ex.Message.ToString();
                return sErr;
            }
            finally
            {
            }
            return string.Empty;
        }
        #endregion
    }
}