using System;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using System.Windows.Forms;

namespace NelaFile
{
    public class BtwType
    {
        public const string KTN = "KTN";
        public const string STN = "STN";
        public const string SIGNO = "SIG";
    }
    public class NelaFile_MachineName
    {
        public const string Nela1 = "Nela 1";
        public const string Nela2 = "Nela 2";
        public const string Nela3 = "Nela 3";
        public const string Nela4 = "Nela 4";
        public const string Estar = "TES3.0"; //"ESTAR";
        public const string E3PR = "E3PR";
        public const string DA001 = "DA001";
        public const string Elektra = "Elektra";
    }
    public class TNS_Names
    {
        public const string AD9PROD = "AD9PROD";
        public const string AD9TEST = "AD9TEST";
        public const string AD9UPG3 = "AD9UPG3";
        public const string ADPROD = "ADPROD";
    }
    public class IFS_Versions
    {
        public const string IFS_2003 = "IFS_2003";
        public const string IFS_9 = "IFS_9";
    }

    class Globals
    {

        public static string TNSNAME = TNS_Names.AD9PROD;
        public static string IFS_VERSION = IFS_Versions.IFS_9;
        public static string ConString = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456";
        public static string ConStrin_AD9PROD__Elektra = "Data Source=" + Globals.TNSNAME + ";User Id=elektra;Password=eLektra2020*";
        //public static string ConString_IFS9 = "Data Source=" + Globals.TNSNAME + ";User Id=lotyl;Password=123456";
        public static string PRINTER_DEFAULT = string.Empty;
        public static string PRINTER = string.Empty;
        //public static string TNSNAME = "AD9UPG3";
        public static string MUJ_EMAIL = "miroslav.senkerik@austin.cz";
        public static string NOREPLY_EMAIL = "noreply@austin.cz";
        public static string EMAIL_FROM = "From@" + Application.ProductName.ToString() + ".cz";
        public static string EMAIL_BODY = string.Empty;
        public static string EMAIL_SUBJECT = string.Empty;
        public static string SMTP_SERVER = "relay.austin.cz";
        public static string IMT_FTP_SERVER = "192.168.61.15";
        public static string IMT_FTP_SERVER_IMT = "192.168.61.11";
        public static string IMT_FILE_RECEPTURA = string.Empty;
        public static int  ID_TABLE_NELA = 0;
        public static string IMT_FOLDERFILE_RECEPTURA = string.Empty;
        public static string FOLDER_BAP01_NALEPKY = "\\bap01\\IFS_NALEPKY\\";
        public static string APPLICATION_NAME = Application.ProductName + ".EXE";
        public static string FOLDER_NALEPKY = "\\\\Bap01\\IFS_NALEPKY\\";
        public static string FOLDER_NEW_VERSION = "\\\\austin.cz\\cz\\AD\\PROGRAM\\NelaFile\\";
        public static string FOLDER_OUTPUT = "\\\\austin.cz\\cz\\AD\\PROGRAM\\NelaFile\\output\\";
        public static string FOLDER_BTW_EXPORT = string.Empty;
        public static string ConString_AUSTIN_SLC = "Persist Security Info=True;User ID=sl;Password=123456;Initial Catalog=AUSTIN_SLC;Data Source=SQL01";
        public static string ConString_MSSQL_Austin_Trc = "Persist Security Info=True;User ID=trc;Password=123456;Initial Catalog=AUSTIN_TRC;Data Source=sql01";

        // vendor
        public static string SQL_Get_ProducerCode_Prefix = "SELECT a.c_ex_tracer_prefix FROM IFSAPP.COMPANY a where a.company = 'AD'";
        public static string SQL_Get_xml_production_site_number = "select a.value from IFSAPP.COMM_METHOD a where a.party_type_db  = 'COMPANY' AND a.identity = 'AD' AND a.method_id = 'Interkom' AND a.name = 'xml_production_site_number'";
        public static string SQL_GetVendor_Company = "SELECT a.contract FROM IFSAPP.CUSTOMER_ORDER a where a.order_no = 'xxORDER_NO'";
        public static string SQL_GetVendor_Address_ID = "SELECT a.contract FROM IFSAPP.CUSTOMER_ORDER a where a.order_no = 'xxORDER_NO'";

        public static string SQL_GetVendor_Name = "SELECT a.name FROM IFSAPP.a0000_company_ad a where a.def_lang = 'en'";
        public static string SQL_Get_Language_Code_FromOrder = "SELECT a.language_code FROM IFSAPP.customer_order a where a.order_no = 'xxORDER_NO'";
        public static string SQL_GetVendor_Country = "SELECT a.country FROM IFSAPP.a0000_company_ad a where a.def_lang = 'xxLANGUAGE_CODE'";
        public static string SQL_GetVendor_Address = "SELECT a.street || ' ' || a.house_no FROM IFSAPP.a0000_company_ad a where a.def_lang = 'en'";
        public static string SQL_GetVendor_City = "SELECT a.city FROM IFSAPP.a0000_company_ad a where a.def_lang = 'en'";
        public static string SQL_GetVendor_ZipCode = "SELECT a.ZIP_CODE FROM IFSAPP.a0000_company_ad a where a.def_lang = 'en'";
        // receiver
        public static string SQL_GetIDZakaznika = "SELECT a.customer_no FROM IFSAPP.CUSTOMER_ORDER a where a.order_no = 'xxORDER_NO'";
        public static string SQL_GetCompanyNameZakaznika = "SELECT a.name FROM IFSAPP.CUSTOMER_INFO a where a.customer_id = 'xxCustomer_id'";
        public static string SQL_GetCustomer_COMM_METHOD = "SELECT xxColumn FROM IFSAPP.CUSTOMER_INFO_COMM_METHOD a where a.customer_id = 'xxCustomer_id' " +
                                                           " AND a.address_id = 'xxAddress_id' AND a.method_default = 'TRUE'";

        public static string SQL_GetCustomer_COMM_METHOD_FAX = "SELECT a.value FROM IFSAPP.CUSTOMER_INFO_COMM_METHOD a where a.customer_id = 'xxCustomer_id' " +
                                                           " AND a.address_id = 'xxAddress_id' AND a.method_id_db = 'FAX'";

        public static string SQL_GetCustomer_COMM_METHOD_EMAIL = "SELECT a.value FROM IFSAPP.CUSTOMER_INFO_COMM_METHOD a where a.customer_id = 'xxCustomer_id' " +
                                                           " AND a.address_id = 'xxAddress_id' AND a.method_id_db = 'E_MAIL'";

        public static string SQL_GetCustomer_COMM_METHOD_MOBILE = "SELECT a.value FROM IFSAPP.CUSTOMER_INFO_COMM_METHOD a where a.customer_id = 'xxCustomer_id' " +
                                                           " AND a.address_id = 'xxAddress_id' AND a.method_id_db = 'MOBILE'";

        public static string SQL_GetIDAddress = "SELECT a.SHIP_ADDR_NO FROM IFSAPP.CUSTOMER_ORDER a where a.order_no = 'xxORDER_NO'";
        public static string SQL_GetAddress = "SELECT xxColumn  FROM IFSAPP.CUSTOMER_INFO_ADDRESS a where a.customer_id = 'xxCustomer_id' AND a.address_id = 'xxAddress_id'";
        public static string SQL_VyberZeSady = "SELECT (a.sub_doc_class || '-' || a.sub_doc_no || '-' || a.sub_doc_rev || '.' || b.format_size) as FileName," +
            "IFSAPP.DOC_TITLE_API.Get_Title(a.sub_doc_class, a.SUB_DOC_NO) as Title" +
            //" b.doc_class,b.doc_no,b.doc_rev" +
            " FROM IFSAPP.DOC_STRUCTURE a, IFSAPP.DOC_ISSUE_REFERENCE b" +
            " where a.sub_doc_class = b.doc_class AND a.sub_doc_no = b.doc_no AND a.sub_doc_rev = b.doc_rev" +
            " AND a.doc_no IN " +
            "(" +
            "  SELECT" +
            "  a.doc_no" +
            "  FROM IFSAPP.DOC_REFERENCE_OBJECT a" +
            "  WHERE 1 = 1" +
            "  AND substr(a.key_value, 1, INSTR(a.key_value, '^', 1, 1) - 1) = 'xxORDER_NO'" +
            "  AND substr(a.key_value, INSTR(a.key_value, '^', 1, 1) + 1, INSTR(a.key_value, '^', 1, 2) - INSTR(a.key_value, '^', 1, 1) - 1) = 'xxLINE_NO'" +
            "  AND substr(a.key_value, INSTR(a.key_value, '^', 1, 2) + 1, INSTR(a.key_value, '^', 1, 3) - INSTR(a.key_value, '^', 1, 2) - 1) =  'xxRELEASE_NO'" +
            //"  AND substr(a.key_value,INSTR(a.key_value, '^', 1, 3)+1,INSTR(a.key_value, '^', 1, 4) - INSTR(a.key_value, '^', 1, 3)-1) =  '" + _LINE_ITEM_NO + "'" +
            "   AND a.doc_class LIKE 'SADA%'" +
            ")" +
            " AND a.sub_doc_class LIKE '%xxTypNalepky%'";

        public static string SQL_DOC_CLASS = "SELECT sub_doc_class as doc_class, sub_doc_no as doc_no,sub_doc_rev as doc_rev " +
            " FROM IFSAPP.DOC_STRUCTURE a, IFSAPP.DOC_ISSUE_REFERENCE b" +
            " where a.sub_doc_class = b.doc_class AND a.sub_doc_no = b.doc_no AND a.sub_doc_rev = b.doc_rev" +
            " AND a.doc_no IN " +
            "(" +
            "  SELECT" +
            "  a.doc_no" +
            "  FROM IFSAPP.DOC_REFERENCE_OBJECT a" +
            "  WHERE 1 = 1" +
            "  AND substr(a.key_value, 1, INSTR(a.key_value, '^', 1, 1) - 1) = 'xxORDER_NO'" +
            "  AND substr(a.key_value, INSTR(a.key_value, '^', 1, 1) + 1, INSTR(a.key_value, '^', 1, 2) - INSTR(a.key_value, '^', 1, 1) - 1) = 'xxLINE_NO'" +
            "  AND substr(a.key_value, INSTR(a.key_value, '^', 1, 2) + 1, INSTR(a.key_value, '^', 1, 3) - INSTR(a.key_value, '^', 1, 2) - 1) =  'xxRELEASE_NO'" +
            //"  AND substr(a.key_value,INSTR(a.key_value, '^', 1, 3)+1,INSTR(a.key_value, '^', 1, 4) - INSTR(a.key_value, '^', 1, 3)-1) =  '" + _LINE_ITEM_NO + "'" +
            //"   AND a.doc_class LIKE 'SADA%'" +
            ")" +
            " AND (b.doc_class || '-' || b.doc_no || '-' || b.doc_rev || '.' || b.format_size) = 'xxFILE_NAME'";
            //" AND a.sub_doc_class LIKE '%xxTypNalepky%'";

        public static string SQL_FileProDA = "" +
            "SELECT doc_rev,CUSTOMER_PO_NO,order_no,line_no,release_no,line_item_no,package_num1,package_num2,catalog_no,catalog_desc," +
            "package_qty, small_package_qty, package_gross_weight, round_net_weight, package_net_weight, explosive_net_weight,un_no,adr_rid_class," +
            "detonator_class,delay_level,cavity_material,ignition_cue_mj,ignition_cue_j,safety_stream,safety_cue_mj," +
            "safety_cue_j,supply_wire_material,wire_isulator_material,wire_length,wire_length_trace," +
            "consumption_date,nominal_time_ea,nominal_time_na,st_length,delay_step,batch_no," +
            "testing_station_number,testing_protocol_number,customer_authorize_decree,another_notice,note,general,general_1," +
            "PocetSkl," +
            "'CZ010' AS ST_PRODUCER," +
            "(SELECT" +
            "   (SELECT characteristic_value_desc FROM IFSAPP.DISCRETE_CHARAC_VALUE WHERE  characteristic_code = t.characteristic_code AND characteristic_value = t.attr_value)" +
            " FROM IFSAPP.inv_part_discrete_char t WHERE(part_no = x.catalog_no)" +
            " AND  t.characteristic_code = 'N008'" +
            " AND t.contract = 'AD') as ST_COLOR," +
            "(SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR ds WHERE  ds.CHARACTERISTIC_CODE = 'NL03' AND ds.CONTRACT = 'AD'" +
            " AND ds.part_no = (SELECT part_no FROM A0055_VyrabenePolozky vp WHERE vp.VZ_BatchNo= pVZBatchNo)) AS STOPPER," +
            "(SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR ds WHERE  ds.CHARACTERISTIC_CODE = 'NL01' AND ds.CONTRACT = 'AD'" +
            " AND ds.part_no = (SELECT part_no FROM A0055_VyrabenePolozky vp WHERE vp.VZ_BatchNo = pVZBatchNo)) AS CONNECTOR," +
            "(SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR ds WHERE  ds.CHARACTERISTIC_CODE = 'NL02' AND ds.CONTRACT = 'AD'" +
            " AND ds.part_no = (SELECT part_no FROM A0055_VyrabenePolozky vp WHERE vp.VZ_BatchNo = pVZBatchNo)) AS SEALING," +
            "(SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR ds WHERE  ds.CHARACTERISTIC_CODE = 'N017' AND ds.CONTRACT = 'AD'" +
            " AND ds.part_no = x.catalog_no) AS DETONATOR_LENGTH," +
           "(SELECT" +
            "  (SELECT characteristic_value_desc FROM IFSAPP.DISCRETE_CHARAC_VALUE WHERE  characteristic_code = t.characteristic_code AND characteristic_value = t.attr_value)" +
            "FROM IFSAPP.inv_part_discrete_char t WHERE(part_no = x.catalog_no)" +
            "AND  t.characteristic_code = 'NL04'" +
            "AND t.contract = 'AD') as FLAG_COLOR," +
            "(SELECT attr_value FROM IFSAPP.INV_PART_DISCRETE_CHAR ds WHERE  ds.CHARACTERISTIC_CODE = 'NL05' AND ds.CONTRACT = 'AD'" +
            " AND ds.part_no = (SELECT part_no FROM A0055_VyrabenePolozky vp WHERE vp.VZ_BatchNo= pVZBatchNo))" +
            "AS TN_COLOR," +
            "Od AS PACKAGE_NUM2BEG, Do AS PACKAGE_NUM2END, PocetSkl AS SEQ_TOTAL, KsVKartonu," +
            "(Do - Od + 1) PocetKartonu," +
            "(SELECT SUM(cbs.sub_pack_amount) from C_BARCODE_SHPORD cbs" +
            " WHERE cbs.customer_order = pOrder_no" +
            " AND cbs.customer_line_no = pLine_no" +
            " AND cbs.customer_rel_no = pRelease_no" +
            " AND cbs.customer_line_item_no = pLine_item_no)  AS TOTAL_PCS," +
            "zo_ks as Kontrolni_soucet,TO_CHAR(manufacturedate,'YYMMDD') AS manufacture_date," +
            "TO_CHAR(add_months(manufacturedate, 24), 'YYMMDD')  AS expiration_date" +
            " FROM " +
            "(" +
                "SELECT batch_no," +
                " MAX((select NVL(PART_NO_FOR_CUSTOMER, catalog_no) from CUSTOMER_ORDER_LINE WHERE order_no = pOrder_no AND line_no = pLine_no AND rel_no = pRelease_no AND line_item_no = pLine_item_no)) AS catalog_no," +
                " MAX(a.catalog_desc)catalog_desc," +
                " MIN(package_num2) AS Od, MAX(package_num2) AS Do, MAX(a.seq_num) AS PocetSkl, MAX(package_qty) AS KsVKartonu," +
                " MAX(small_package_qty) AS KsVeSkladacce," +
                " MAX((SELECT MAX(col.DESIRED_QTY) FROM IFSAPP.customer_order_line col WHERE col.order_no = a.order_no AND col.line_no = a.line_no AND col.rel_no = a.release_no)) AS ZO_Ks," +
                " MAX(a.doc_rev)doc_rev, MAX(a.customer_po_no) customer_po_no, MAX(a.order_no) order_no," +
                " MAX(a.line_no) line_no, MAX(a.release_no) release_no, MAX(a.line_item_no) line_item_no," +
                " MAX(a.package_num1) package_num1, MAX(a.package_num2) package_num2," +
                " MAX(a.package_qty) package_qty, MAX(a.small_package_qty) small_package_qty," +
                " MAX(a.package_gross_weight) package_gross_weight, MAX(a.round_net_weight) round_net_weight," +
                " MAX(a.package_net_weight) package_net_weight, MAX(a.explosive_net_weight) explosive_net_weight," +
                " MAX(a.un_no) un_no, MAX(a.adr_rid_class) adr_rid_class, MAX(a.detonator_class) detonator_class," +
                " MAX(a.delay_level) delay_level, MAX(a.cavity_material) cavity_material, MAX(a.ignition_cue_mj) ignition_cue_mj," +
                " MAX(a.ignition_cue_j) ignition_cue_j, MAX(a.safety_stream) safety_stream, MAX(a.safety_cue_mj) safety_cue_mj," +
                " MAX(a.safety_cue_j) safety_cue_j, MAX(a.supply_wire_material) supply_wire_material," +
                " MAX(a.wire_isulator_material) wire_isulator_material," +
                " MAX(a.wire_length) wire_length," +
                " MAX(a.wire_length_trace) wire_length_trace, MAX(a.consumption_date) consumption_date, MAX(a.nominal_time_ea) nominal_time_ea," +
                " MAX(a.nominal_time_na) nominal_time_na, MAX(a.st_lenght) st_length, MAX(a.delay_step) delay_step," +
                " MAX(a.testing_station_number) testing_station_number," +
                " MAX(a.testing_protocol_number) testing_protocol_number," +
                " MAX(a.customer_authorize_decree) customer_authorize_decree, MAX(a.another_notice) another_notice, MAX(a.note) note," +
                " MAX(a.general) general, MAX(a.general_1) general_1," +
                " MAX((SELECT MAX(col.wanted_delivery_date) FROM IFSAPP.customer_order_line col" +
                "      WHERE col.order_no = a.order_no AND col.line_no = a.line_no AND col.rel_no = a.release_no))  AS manufacturedate" +
                "FROM IFSAPP.c_labels_ad a" +
                "WHERE a.order_no = pOrder_no" +
                "AND a.line_no = pLine_no" +
                "AND a.release_no = pRelease_no" +
                "AND a.line_item_no = pLine_item_no" +
                "GROUP BY a.batch_no ORDER BY a.batch_no" +
      ") x";
        public static string SQL_Nalepky_From_PO = "SELECT (c.doc_class || '-' || c.doc_no || '-' ||c.doc_rev || '.' || c.format_size) as FILE_NAME,c.doc_class" +
              " FROM IFSAPP.DOC_REFERENCE_OBJECT a, IFSAPP.DOC_STRUCTURE b,IFSAPP.DOC_ISSUE_REFERENCE c" +
              " WHERE a.doc_class = b.doc_class" +
              " AND a.doc_no = b.doc_no" +
              " AND a.doc_rev = b.doc_rev" +
              " AND b.sub_doc_class = c.doc_class" +
              " AND b.sub_doc_no  = c.doc_no" +
              " AND b.sub_doc_rev = c.doc_rev" +
              " AND a.lu_name = 'CustomerOrderLine'" +
              " AND a.key_value LIKE  '";

        public static string SQL_Nalepky_From_DMS_IFS9 = "SELECT d.file_name,c.doc_class" +
              " FROM IFSAPP.DOC_REFERENCE_OBJECT a, IFSAPP.DOC_STRUCTURE b,IFSAPP.DOC_ISSUE_REFERENCE c,IFSAPP.EDM_FILE d " +
              " WHERE a.doc_class = b.doc_class" +
              " AND a.doc_no = b.doc_no" +
              " AND a.doc_rev = b.doc_rev" +
              " AND b.sub_doc_class = c.doc_class" +
              " AND b.sub_doc_no  = c.doc_no" +
              " AND b.sub_doc_rev = c.doc_rev" +
              " AND c.doc_class = d.doc_class AND c.doc_no = d.doc_no" +
              " AND a.lu_name = 'CustomerOrderLine'" +
              " AND a.key_value LIKE  '";
        public static string JPG_LAST_FOLDER_FILENAME = string.Empty;
        public static string NELAFILE_LogFile = Application.ProductName.ToString() + ".log";
        public static string DECIMAL_SEPARATOR = DecimalSeparator.Comma;
        public static string FOLDER_NALEPKY_IFS9PROD = @"\\austin.cz\cz\AD\SHARE\IT_IFS_DOC_MAN\AD9PROD\NALEPKY\";
        public static string FOLDER_NALEPKY_IFS9TEST = @"\\austin.cz\cz\AD\SHARE\IT_IFS_DOC_MAN\AD9TEST\NALEPKY\";
        public static string FOLDER_E3PR_DEFAULT = @"\\austin.cz\cz\AD\SHARE\UO_business\Estar\";
        public static string FILE_E3PR_TXT = string.Empty;
        public static string FOLDER_DA001_DEFAULT = @"\\austin.cz\cz\AD\SHARE\UO_business\Estar\";
        public static string FILE_DA001_TXT = string.Empty;
    }
    public class ZO
    {
        // udaje o zakazce
        public static string ORDER_NO = string.Empty;
        public static string LINE_NO = string.Empty;
        public static string RELEASE_NO = string.Empty;
        public static string LINE_ITEM_NO = string.Empty;

        public static string ZO_BatchNo = string.Empty;
        public static string VZ_BatchNo = string.Empty;
        
        public static string PART_NO = string.Empty;
        public static string Customer = string.Empty;
        public static string Address = string.Empty;

    }
    public class DecimalSeparator
    {
        public const string Comma = ",";
        public const string Dot = ".";
    }
    public class btw
    {
        public const string BTW_TYPE = "BTW_TYPE";


        public const string DOC_CLASS = "c01_DOC_CLASS";
        public const string TITLE = "c02_TITLE";
        public const string DOC_NO = "c03_DOC_NO";
        public const string DOC_REV = "c04_DOC_REV";
        public const string FILE_NAME = "c05_FILE_NAME";
        public const string CUSTOMER_PO_NO = "c06_CUSTOMER_PO_NO";
        public const string ORDER_NO = "c07_ORDER_NO";
        public const string LINE_NO = "c08_LINE_NO";
        public const string RELEASE_NO = "c09_RELEASE_NO";
        public const string LINE_ITEM_NO = "c10_LINE_ITEM_NO";
        public const string PACKAGE_NUM1 = "c11_PACKAGE_NUM1";
        public const string SEQ_NUM = "c12_SEQ_NUM";
        public const string PACKAGE_NUM2 = "c13_PACKAGE_NUM2";
        public const string CATALOG_NO = "c14_CATALOG_NO";
        public const string CATALOG_DESC = "c15_CATALOG_DESC";
        public const string VENDOR_COMPANY = "c16_VENDOR_COMPANY";
        public const string VENDOR_NAME = "c17_VENDOR_NAME";
        public const string VENDOR_ADDRESS_ID = "c18_VENDOR_ADDRESS_ID";
        public const string VENDOR_STREET = "c19_VENDOR_STREET";
        public const string VENDOR_CITY = "c20_VENDOR_CITY";
        public const string VENDOR_ZIP_CODE = "c21_VENDOR_ZIP_CODE";
        public const string VENDOR_COUNTRY = "c22_VENDOR_COUNTRY";
        public const string RECEIVER_COMPANY_NAME2 = "c23_RECEIVER_COMPANY_NAME2";
        public const string RECEIVER_ADDRESS1 = "c24_RECEIVER_ADDRESS1";
        public const string RECEIVER_ADDRESS2 = "c25_RECEIVER_ADDRESS2";
        public const string RECEIVER_ZIP_CODE = "c26_RECEIVER_ZIP_CODE";
        public const string RECEIVER_CITY = "c27_RECEIVER_CITY";
        public const string RECEIVER_COUNTRY = "c28_RECEIVER_COUNTRY";
        public const string RECEIVER_NAME = "c29_RECEIVER_NAME";
        public const string RECEIVER_PHONE = "c30_RECEIVER_PHONE";
        public const string RECEIVER_FAX = "c31_RECEIVER_FAX";
        public const string RECEIVER_MOBILE = "c32_RECEIVER_MOBILE";
        public const string RECEIVER_EMAIL = "c33_RECEIVER_EMAIL";
        public const string UNIQUE_PACKAGE_ID = "c34_UNIQUE_PACKAGE_ID";
        public const string UNIQUE_SUBPACKAGE_ID = "c35_UNIQUE_SUBPACKAGE_ID";
        public const string PACKAGE_QTY = "c36_PACKAGE_QTY";
        public const string SMALL_PACKAGE_QTY = "c37_SMALL_PACKAGE_QTY";
        public const string PACKAGE_GROSS_WEIGHT = "c38_PACKAGE_GROSS_WEIGHT";
        public const string ROUND_NET_WEIGHT = "c39_ROUND_NET_WEIGHT";
        public const string PACKAGE_NET_WEIGHT = "c40_PACKAGE_NET_WEIGHT";
        public const string EXPLOSIVE_NET_WEIGHT = "c41_EXPLOSIVE_NET_WEIGHT";
        public const string UN_NO = "c42_UN_NO";
        public const string ADR_RID_CLASS = "c42_UN_NO";
        public const string DETONATOR_CLASS = "c44_DETONATOR_CLASS";
        public const string DELAY_LEVEL = "c45_DELAY_LEVEL";
        public const string CAVITY_MATERIAL = "c46_CAVITY_MATERIAL";
        public const string IGNITION_CUE_MJ = "c47_IGNITION_CUE_MJ";
        public const string IGNITION_CUE_J = "c48_IGNITION_CUE_J";
        public const string SAFETY_STREAM = "c49_SAFETY_STREAM";
        public const string SAFETY_CUE_MJ = "c50_SAFETY_CUE_MJ";
        public const string SAFETY_CUE_J = "c51_SAFETY_CUE_J";
        public const string SUPPLY_WIRE_MATERIAL = "c52_SUPPLY_WIRE_MATERIAL";
        public const string WIRE_ISULATOR_MATERIAL = "c53_WIRE_ISULATOR_MATERIAL";
        public const string WIRE_LENGTH = "c54_WIRE_LENGTH";
        public const string WIRE_LENGTH_TRACE = "c55_WIRE_LENGTH_TRACE";
        public const string CONSUMPTION_DATE = "c56_CONSUMPTION_DATE";
        public const string NOMINAL_TIME_EA = "c57_NOMINAL_TIME_EA";
        public const string NOMINAL_TIME_NA = "c58_NOMINAL_TIME_NA";
        public const string ST_LENGTH = "c59_ST_LENGTH";
        public const string DELAY_STEP = "c60_DELAY_STEP";
        public const string BATCH_NO = "c61_BATCH_NO";
        public const string TESTING_STATION_NUMBER = "c62_TESTING_STATION_NUMBER";
        public const string TESTING_PROTOCOL_NUMBER = "c63_TESTING_PROTOCOL_NUMBER";
        public const string CUSTOMER_AUTHORIZE_DECREE = "c64_CUSTOMER_AUTHORIZE_DECREE";
        public const string ANOTHER_NOTICE = "c65_ANOTHER_NOTICE";
        public const string NOTE = "c66_NOTE";
        public const string OBJVERSION = "c67_OBJVERSION";
        public const string OBJID = "c68_OBJID";
        public const string DATE = "c69_DATE";
        public const string EXPIRATION = "c70_EXPIRATION";
        public const string TOTAL_RESIST = "c71_TOTAL_RESIST";
        public const string BRIDGE_RESIST = "c72_BRIDGE_RESIST";
        public const string DATE_CODE = "c73_DATE_CODE";
        public const string LOT_NUMBER = "c74_LOT_NUMBER";
        public const string STATION_NUM = "c75_STATION_NUM";
        public const string WORKER1 = "c76_WORKERS";
        public const string WORKER2 = "c77_WORKER2";
        public const string DUPLICATE = "c78_DUPLICATE";
        public const string PIECE_NUM = "c79_PIECE_NUM";
        public const string DELAY_FRACTION = "c80_DELAY_FRACTION";
        public const string GENERAL = "c81_GENERAL";
        public const string GENERAL_1 = "c82_GENERAL_1";
        public const string BOX_PCS = "c83_BOX_PCS";
        public const string MANUAL = "c84_MANUAL";
        public const string PRODUCER = "c85_PRODUCER";
        public const string DMX_PACKAGE_QTY = "c86_DMX_PACKAGE_QTY";
        public const string DMX_DELAY_LEVEL = "c87_DMX_DELAY_LEVEL";
        public const string DMX_DELAY_STEP = "c88_DMX_DELAY_STEP";
        public const string DMX_NOMINAL_TIME_EA = "c89_DMX_NOMINAL_TIME_EA";
        public const string DMX_NOMINAL_TIME_NA = "c90_DMX_NOMINAL_TIME_NA";
        public const string DMX_WIRE_LENGTH = "c91_DMX_WIRE_LENGTH";
        public const string DMX_ST_LENGTH = "c92_DMX_ST_LENGTH";
        public const string DMX_DATE = "c93_DMX_DATE";
        public const string DMX_PACKAGE_NUM2 = "c94_DMX_PACKAGE_NUM2";
        public const string DMX_SEQ_NUM = "c95_DMX_SEQ_NUM";
        public const string DMX_PRODUCER = "c96_DMX_PRODUCER";
        public const string AUTO_LABEL = "c97_AUTO_LABEL";
        public const string TOTAL_PACKAGE_QTY = "c98_TOTAL_PACKAGE_QTY";

        public const string c01_DOC_CLASS = "c01_DOC_CLASS";
        public const string c02_TITLE = "c02_TITLE";
        public const string c03_DOC_NO = "c03_DOC_NO";
        public const string c04_DOC_REV = "c04_DOC_REV";
        public const string c05_FILE_NAME = "c05_FILE_NAME";
        public const string c06_CUSTOMER_PO_NO = "c06_CUSTOMER_PO_NO";
        public const string c07_ORDER_NO = "c07_ORDER_NO";
        public const string c08_LINE_NO = "c08_LINE_NO";
        public const string c09_RELEASE_NO = "c09_RELEASE_NO";
        public const string c10_LINE_ITEM_NO = "c10_LINE_ITEM_NO";
        public const string c11_PACKAGE_NUM1 = "c11_PACKAGE_NUM1";
        public const string c12_SEQ_NUM = "c12_SEQ_NUM";
        public const string c13_PACKAGE_NUM2 = "c13_PACKAGE_NUM2";
        public const string c14_CATALOG_NO = "c14_CATALOG_NO";
        public const string c15_CATALOG_DESC = "c15_CATALOG_DESC";
        public const string c16_VENDOR_COMPANY = "c16_VENDOR_COMPANY";
        public const string c17_VENDOR_NAME = "c17_VENDOR_NAME";
        public const string c18_VENDOR_ADDRESS_ID = "c18_VENDOR_ADDRESS_ID";
        public const string c19_VENDOR_STREET = "c19_VENDOR_STREET";
        public const string c20_VENDOR_CITY = "c20_VENDOR_CITY";
        public const string c21_VENDOR_ZIP_CODE = "c21_VENDOR_ZIP_CODE";
        public const string c22_VENDOR_COUNTRY = "c22_VENDOR_COUNTRY";
        public const string c23_RECEIVER_COMPANY_NAME2 = "c23_RECEIVER_COMPANY_NAME2";
        public const string c24_RECEIVER_ADDRESS1 = "c24_RECEIVER_ADDRESS1";
        public const string c25_RECEIVER_ADDRESS2 = "c25_RECEIVER_ADDRESS2";
        public const string c26_RECEIVER_ZIP_CODE = "c26_RECEIVER_ZIP_CODE";
        public const string c27_RECEIVER_CITY = "c27_RECEIVER_CITY";
        public const string c28_RECEIVER_COUNTRY = "c28_RECEIVER_COUNTRY";
        public const string c29_RECEIVER_NAME = "c29_RECEIVER_NAME";
        public const string c30_RECEIVER_PHONE = "c30_RECEIVER_PHONE";
        public const string c31_RECEIVER_FAX = "c31_RECEIVER_FAX";
        public const string c32_RECEIVER_MOBILE = "c32_RECEIVER_MOBILE";
        public const string c33_RECEIVER_EMAIL = "c33_RECEIVER_EMAIL";
        public const string c34_UNIQUE_PACKAGE_ID = "c34_UNIQUE_PACKAGE_ID";
        public const string c35_UNIQUE_SUBPACKAGE_ID = "c35_UNIQUE_SUBPACKAGE_ID";
        public const string c36_PACKAGE_QTY = "c36_PACKAGE_QTY";
        public const string c37_SMALL_PACKAGE_QTY = "c37_SMALL_PACKAGE_QTY";
        public const string c38_PACKAGE_GROSS_WEIGHT = "c38_PACKAGE_GROSS_WEIGHT";
        public const string c39_ROUND_NET_WEIGHT = "c39_ROUND_NET_WEIGHT";
        public const string c40_PACKAGE_NET_WEIGHT = "c40_PACKAGE_NET_WEIGHT";
        public const string c41_EXPLOSIVE_NET_WEIGHT = "c41_EXPLOSIVE_NET_WEIGHT";
        public const string c42_UN_NO = "c42_UN_NO";
        public const string c43_ADR_RID_CLASS = "c42_UN_NO";
        public const string c44_DETONATOR_CLASS = "c44_DETONATOR_CLASS";
        public const string c45_DELAY_LEVEL = "c45_DELAY_LEVEL";
        public const string c46_CAVITY_MATERIAL = "c46_CAVITY_MATERIAL";
        public const string c47_IGNITION_CUE_MJ = "c47_IGNITION_CUE_MJ";
        public const string c48_IGNITION_CUE_J = "c48_IGNITION_CUE_J";
        public const string c49_SAFETY_STREAM = "c49_SAFETY_STREAM";
        public const string c50_SAFETY_CUE_MJ = "c50_SAFETY_CUE_MJ";
        public const string c51_SAFETY_CUE_J = "c51_SAFETY_CUE_J";
        public const string c52_SUPPLY_WIRE_MATERIAL = "c52_SUPPLY_WIRE_MATERIAL";
        public const string c53_WIRE_ISULATOR_MATERIAL = "c53_WIRE_ISULATOR_MATERIAL";
        public const string c54_WIRE_LENGTH = "c54_WIRE_LENGTH";
        public const string c55_WIRE_LENGTH_TRACE = "c55_WIRE_LENGTH_TRACE";
        public const string c56_CONSUMPTION_DATE = "c56_CONSUMPTION_DATE";
        public const string c57_NOMINAL_TIME_EA = "c57_NOMINAL_TIME_EA";
        public const string c58_NOMINAL_TIME_NA = "c58_NOMINAL_TIME_NA";
        public const string c59_ST_LENGTH = "c59_ST_LENGTH";
        public const string c60_DELAY_STEP = "c60_DELAY_STEP";
        public const string c61_BATCH_NO = "c61_BATCH_NO";
        public const string c62_TESTING_STATION_NUMBER = "c62_TESTING_STATION_NUMBER";
        public const string c63_TESTING_PROTOCOL_NUMBER = "c63_TESTING_PROTOCOL_NUMBER";
        public const string c64_CUSTOMER_AUTHORIZE_DECREE = "c64_CUSTOMER_AUTHORIZE_DECREE";
        public const string c65_ANOTHER_NOTICE = "c65_ANOTHER_NOTICE";
        public const string c66_NOTE = "c66_NOTE";
        public const string c67_OBJVERSION = "c67_OBJVERSION";
        public const string c68_OBJID = "c68_OBJID";
        public const string c69_DATE = "c69_DATE";
        public const string c70_EXPIRATION = "c70_EXPIRATION";
        public const string c71_TOTAL_RESIST = "c71_TOTAL_RESIST";
        public const string c72_BRIDGE_RESIST = "c72_BRIDGE_RESIST";
        public const string c73_DATE_CODE = "c73_DATE_CODE";
        public const string c74_LOT_NUMBER = "c74_LOT_NUMBER";
        public const string c75_STATION_NUM = "c75_STATION_NUM";
        public const string c76_WORKERS = "c76_WORKERS";
        public const string c77_WORKER2 = "c77_WORKER2";
        public const string c78_DUPLICATE = "c78_DUPLICATE";
        public const string c79_PIECE_NUM = "c79_PIECE_NUM";
        public const string c80_DELAY_FRACTION = "c80_DELAY_FRACTION";
        public const string c81_GENERAL = "c81_GENERAL";
        public const string c82_GENERAL_1 = "c82_GENERAL_1";
        public const string c83_BOX_PCS = "c83_BOX_PCS";
        public const string c84_MANUAL = "c84_MANUAL";
        public const string c85_PRODUCER = "c85_PRODUCER";
        public const string c86_DMX_PACKAGE_QTY = "c86_DMX_PACKAGE_QTY";
        public const string c87_DMX_DELAY_LEVEL = "c87_DMX_DELAY_LEVEL";
        public const string c88_DMX_DELAY_STEP = "c88_DMX_DELAY_STEP";
        public const string c89_DMX_NOMINAL_TIME_EA = "c89_DMX_NOMINAL_TIME_EA";
        public const string c90_DMX_NOMINAL_TIME_NA = "c90_DMX_NOMINAL_TIME_NA";
        public const string c91_DMX_WIRE_LENGTH = "c91_DMX_WIRE_LENGTH";
        public const string c92_DMX_ST_LENGTH = "c92_DMX_ST_LENGTH";
        public const string c93_DMX_DATE = "c93_DMX_DATE";
        public const string c94_DMX_PACKAGE_NUM2 = "c94_DMX_PACKAGE_NUM2";
        public const string c95_DMX_SEQ_NUM = "c95_DMX_SEQ_NUM";
        public const string c96_DMX_PRODUCER = "c96_DMX_PRODUCER";
        public const string c97_AUTO_LABEL = "c97_AUTO_LABEL";
        public const string c98_TOTAL_PACKAGE_QTY = "c98_TOTAL_PACKAGE_QTY";

    }
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.ThreadException += Application_ThreadException;

            try
            {
                string sFileNew = "";
                string sFileOld = "";
                csVerze csV = new csVerze();


                // files


                // \\austin.cz\cz\AD\PROGRAM\NelaFile\OUTPUT

                //string sPathNew = "\\\\HAD01\\InstallAplikaceAD\\ProgramyAD\\Nela\\";

                string sPathNew = Globals.FOLDER_NEW_VERSION; // "\\\\austin.cz\\cz\\AD\\PROGRAM\\NelaFile\\";
                sFileNew = sPathNew + Globals.APPLICATION_NAME;
                sFileOld = Application.StartupPath + "\\" + Globals.APPLICATION_NAME;
                string sVerzeOld = "";
                string sVerzeNew = "";


                if (csV.ExistujeNovaVerze(sFileOld, sFileNew, ref sVerzeOld, ref sVerzeNew) == true)
                {
                    if (MessageBox.Show("Nalezena nová verze - provede se kopírování. Spusťte aplikaci znovu.", "Existuje nová verze", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        // main EXE
                        MessageBox.Show("Zkopíruji, stiskněte OK." + Environment.NewLine + "po chvíli pusťte program znovu.","Nová verze",MessageBoxButtons.OK,MessageBoxIcon.Information );
                        //csV.SendEmail("Verze", "Stará verze :" + sVerzeOld + "\r\n" + "Nová verze : " + sVerzeNew,Globals.SMTP_SERVER);
                        csV.cmdCopyFile(sFileOld, sFileNew);
                        Application.Exit();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                //Application.Exit();
                //return;
            }


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message);
            Application.Exit();
        }

    }
}
