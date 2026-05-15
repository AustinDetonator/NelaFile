using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NelaFile
{
    class Database
    {
    }
    public class BtwHdr
    {

        public const string BTW_TYPE = "BTW_TYPE";
        public const string c01_DOC_CLASS = "01_DOC_CLASS";
        public const string c02_TITLE = "02_TITLE";
        public const string c03_DOC_NO = "03_DOC_NO";
        public const string c04_DOC_REV = "04_DOC_REV";
        public const string c05_FILE_NAME = "05_FILE_NAME";
        public const string c06_CUSTOMER_PO_NO = "06_CUSTOMER_PO_NO";
        public const string c07_ORDER_NO = "07_ORDER_NO";
        public const string c08_LINE_NO = "08_LINE_NO";
        public const string c09_RELEASE_NO = "09_RELEASE_NO";
        public const string c10_LINE_ITEM_NO = "10_LINE_ITEM_NO";
        public const string c11_PACKAGE_NUM1 = "11_PACKAGE_NUM1";
        public const string c12_SEQ_NUM = "12_SEQ_NUM";
        public const string c13_PACKAGE_NUM2 = "13_PACKAGE_NUM2";
        public const string c14_CATALOG_NO = "14_CATALOG_NO";
        public const string c15_CATALOG_DESC = "15_CATALOG_DESC";
        public const string c16_VENDOR_COMPANY = "16_VENDOR_COMPANY";
        public const string c17_VENDOR_NAME = "17_VENDOR_NAME";
        public const string c18_VENDOR_ADDRESS_ID = "18_VENDOR_ADDRESS_ID";
        public const string c19_VENDOR_STREET = "19_VENDOR_STREET";
        public const string c20_VENDOR_CITY = "20_VENDOR_CITY";
        public const string c21_VENDOR_ZIP_CODE = "21_VENDOR_ZIP_CODE";
        public const string c22_VENDOR_COUNTRY = "22_VENDOR_COUNTRY";
        public const string c23_RECEIVER_COMPANY_NAME2 = "23_RECEIVER_COMPANY_NAME2";
        public const string c24_RECEIVER_ADDRESS1 = "24_RECEIVER_ADDRESS1";
        public const string c25_RECEIVER_ADDRESS2 = "25_RECEIVER_ADDRESS2";
        public const string c26_RECEIVER_ZIP_CODE = "26_RECEIVER_ZIP_CODE";
        public const string c27_RECEIVER_CITY = "27_RECEIVER_CITY";
        public const string c28_RECEIVER_COUNTRY = "28_RECEIVER_COUNTRY";
        public const string c29_RECEIVER_NAME = "29_RECEIVER_NAME";
        public const string c30_RECEIVER_PHONE = "30_RECEIVER_PHONE";
        public const string c31_RECEIVER_FAX = "31_RECEIVER_FAX";
        public const string c32_RECEIVER_MOBILE = "32_RECEIVER_MOBILE";
        public const string c33_RECEIVER_EMAIL = "33_RECEIVER_EMAIL";
        public const string c34_UNIQUE_PACKAGE_ID = "34_UNIQUE_PACKAGE_ID";
        public const string c35_UNIQUE_SUBPACKAGE_ID = "35_UNIQUE_SUBPACKAGE_ID";
        public const string c36_PACKAGE_QTY = "36_PACKAGE_QTY";
        public const string c37_SMALL_PACKAGE_QTY = "37_SMALL_PACKAGE_QTY";
        public const string c38_PACKAGE_GROSS_WEIGHT = "38_PACKAGE_GROSS_WEIGHT";
        public const string c39_ROUND_NET_WEIGHT = "39_ROUND_NET_WEIGHT";
        public const string c40_PACKAGE_NET_WEIGHT = "40_PACKAGE_NET_WEIGHT";
        public const string c41_EXPLOSIVE_NET_WEIGHT = "41_EXPLOSIVE_NET_WEIGHT";
        public const string c42_UN_NO = "42_UN_NO";
        public const string c43_ADR_RID_CLASS = "43_ADR_RID_CLASS";
        public const string c44_DETONATOR_CLASS = "44_DETONATOR_CLASS";
        public const string c45_DELAY_LEVEL = "45_DELAY_LEVEL";
        public const string c46_CAVITY_MATERIAL = "46_CAVITY_MATERIAL";
        public const string c47_IGNITION_CUE_MJ = "47_IGNITION_CUE_MJ";
        public const string c48_IGNITION_CUE_J = "48_IGNITION_CUE_J";
        public const string c49_SAFETY_STREAM = "49_SAFETY_STREAM";
        public const string c50_SAFETY_CUE_MJ = "50_SAFETY_CUE_MJ";
        public const string c51_SAFETY_CUE_J = "51_SAFETY_CUE_J";
        public const string c52_SUPPLY_WIRE_MATERIAL = "52_SUPPLY_WIRE_MATERIAL";
        public const string c53_WIRE_ISULATOR_MATERIAL = "53_WIRE_ISULATOR_MATERIAL";
        public const string c54_WIRE_LENGTH = "54_WIRE_LENGTH";
        public const string c55_WIRE_LENGTH_TRACE = "55_WIRE_LENGTH_TRACE";
        public const string c56_CONSUMPTION_DATE = "56_CONSUMPTION_DATE";
        public const string c57_NOMINAL_TIME_EA = "57_NOMINAL_TIME_EA";
        public const string c58_NOMINAL_TIME_NA = "58_NOMINAL_TIME_NA";
        public const string c59_ST_LENGTH = "59_ST_LENGTH";
        public const string c60_DELAY_STEP = "60_DELAY_STEP";
        public const string c61_BATCH_NO = "61_BATCH_NO";
        public const string c62_TESTING_STATION_NUMBER = "62_TESTING_STATION_NUMBER";
        public const string c63_TESTING_PROTOCOL_NUMBER = "63_TESTING_PROTOCOL_NUMBER";
        public const string c64_CUSTOMER_AUTHORIZE_DECREE = "64_CUSTOMER_AUTHORIZE_DECREE";
        public const string c65_ANOTHER_NOTICE = "65_ANOTHER_NOTICE";
        public const string c66_NOTE = "66_NOTE";
        public const string c67_OBJVERSION = "67_OBJVERSION";
        public const string c68_OBJID = "68_OBJID";
        public const string c69_DATE = "69_DATE";
        public const string c70_EXPIRATION = "70_EXPIRATION";
        public const string c71_TOTAL_RESIST = "71_TOTAL_RESIST";
        public const string c72_BRIDGE_RESIST = "72_BRIDGE_RESIST";
        public const string c73_DATE_CODE = "73_DATE_CODE";
        public const string c74_LOT_NUMBER = "74_LOT_NUMBER";
        public const string c75_STATION_NUM = "75_STATION_NUM";
        public const string c76_WORKERS = "76_WORKER1";
        public const string c77_WORKER2 = "77_WORKER2";
        public const string c78_DUPLICATE = "78_DUPLICATE";
        public const string c79_PIECE_NUM = "79_PIECE_NUM";
        public const string c80_DELAY_FRACTION = "80_DELAY_FRACTION";
        public const string c81_GENERAL = "81_GENERAL";
        public const string c82_GENERAL_1 = "82_GENERAL_1";
        public const string c83_BOX_PCS = "83_BOX_PCS";
        public const string c84_MANUAL = "84_MANUAL";
        public const string c85_PRODUCER = "85_PRODUCER";
        public const string c86_DMX_PACKAGE_QTY = "86_DMX_PACKAGE_QTY";
        public const string c87_DMX_DELAY_LEVEL = "87_DMX_DELAY_LEVEL";
        public const string c88_DMX_DELAY_STEP = "88_DMX_DELAY_STEP";
        public const string c89_DMX_NOMINAL_TIME_EA = "89_DMX_NOMINAL_TIME_EA";
        public const string c90_DMX_NOMINAL_TIME_NA = "90_DMX_NOMINAL_TIME_NA";
        public const string c91_DMX_WIRE_LENGTH = "91_DMX_WIRE_LENGTH";
        public const string c92_DMX_ST_LENGTH = "92_DMX_ST_LENGTH";
        public const string c93_DMX_DATE = "93_DMX_DATE";
        public const string c94_DMX_PACKAGE_NUM2 = "94_DMX_PACKAGE_NUM2";
        public const string c95_DMX_SEQ_NUM = "95_DMX_SEQ_NUM";
        public const string c96_DMX_PRODUCER = "96_DMX_PRODUCER";
        public const string c97_AUTO_LABEL = "97_AUTO_LABEL";
        public const string c98_TOTAL_PACKAGE_QTY = "98_TOTAL_PACKAGE_QTY";
        public const string c99_UFI_CODE = "99_UFI_CODE";
    }

    public class BtwRow
    {
        public string Document_code { get; set; }

        public string BTW_TYPE { get; set; }
        public string c01_DOC_CLASS { get; set; }
        public string c02_TITLE { get; set; }
        public string c03_DOC_NO { get; set; }
        public string c04_DOC_REV { get; set; } 
        public string c05_FILE_NAME { get; set; } 
        public string c06_CUSTOMER_PO_NO { get; set; } 
        public string c07_ORDER_NO { get; set; } 
        public string c08_LINE_NO { get; set; } 
        public string c09_RELEASE_NO { get; set; } 
        public string c10_LINE_ITEM_NO { get; set; } 
        public string c11_PACKAGE_NUM1 { get; set; } 
        public string c12_SEQ_NUM { get; set; } 
        public string c13_PACKAGE_NUM2 { get; set; } 
        public string c14_CATALOG_NO { get; set; } 
        public string c15_CATALOG_DESC { get; set; } 
        public string c16_VENDOR_COMPANY { get; set; } 
        public string c17_VENDOR_NAME { get; set; } 
        public string c18_VENDOR_ADDRESS_ID { get; set; } 
        public string c19_VENDOR_STREET { get; set; } 
        public string c20_VENDOR_CITY { get; set; } 
        public string c21_VENDOR_ZIP_CODE { get; set; } 
        public string c22_VENDOR_COUNTRY { get; set; } 
        public string c23_RECEIVER_COMPANY_NAME2 { get; set; } 
        public string c24_RECEIVER_ADDRESS1 { get; set; } 
        public string c25_RECEIVER_ADDRESS2 { get; set; } 
        public string c26_RECEIVER_ZIP_CODE { get; set; } 
        public string c27_RECEIVER_CITY { get; set; } 
        public string c28_RECEIVER_COUNTRY { get; set; } 
        public string c29_RECEIVER_NAME { get; set; } 
        public string c30_RECEIVER_PHONE { get; set; } 
        public string c31_RECEIVER_FAX { get; set; } 
        public string c32_RECEIVER_MOBILE { get; set; } 
        public string c33_RECEIVER_EMAIL { get; set; } 
        public string c34_UNIQUE_PACKAGE_ID { get; set; } 
        public string c35_UNIQUE_SUBPACKAGE_ID { get; set; } 
        public string c36_PACKAGE_QTY { get; set; } 
        public string c37_SMALL_PACKAGE_QTY { get; set; } 
        public string c38_PACKAGE_GROSS_WEIGHT { get; set; } 
        public string c39_ROUND_NET_WEIGHT { get; set; } 
        public string c40_PACKAGE_NET_WEIGHT { get; set; } 
        public string c41_EXPLOSIVE_NET_WEIGHT { get; set; } 
        public string c42_UN_NO { get; set; } 
        public string c43_ADR_RID_CLASS { get; set; } 
        public string c44_DETONATOR_CLASS { get; set; } 
        public string c45_DELAY_LEVEL { get; set; } 
        public string c46_CAVITY_MATERIAL { get; set; } 
        public string c47_IGNITION_CUE_MJ { get; set; } 
        public string c48_IGNITION_CUE_J { get; set; } 
        public string c49_SAFETY_STREAM { get; set; } 
        public string c50_SAFETY_CUE_MJ { get; set; } 
        public string c51_SAFETY_CUE_J { get; set; } 
        public string c52_SUPPLY_WIRE_MATERIAL { get; set; } 
        public string c53_WIRE_ISULATOR_MATERIAL { get; set; } 
        public string c54_WIRE_LENGTH { get; set; } 
        public string c55_WIRE_LENGTH_TRACE { get; set; } 
        public string c56_CONSUMPTION_DATE { get; set; } 
        public string c57_NOMINAL_TIME_EA { get; set; } 
        public string c58_NOMINAL_TIME_NA { get; set; } 
        public string c59_ST_LENGTH { get; set; } 
        public string c60_DELAY_STEP { get; set; } 
        public string c61_BATCH_NO { get; set; } 
        public string c62_TESTING_STATION_NUMBER { get; set; } 
        public string c63_TESTING_PROTOCOL_NUMBER { get; set; } 
        public string c64_CUSTOMER_AUTHORIZE_DECREE { get; set; } 
        public string c65_ANOTHER_NOTICE { get; set; } 
        public string c66_NOTE { get; set; } 
        public string c67_OBJVERSION { get; set; } 
        public string c68_OBJID { get; set; } 
        public string c69_DATE { get; set; } 
        public string c70_EXPIRATION { get; set; } 
        public string c71_TOTAL_RESIST { get; set; } 
        public string c72_BRIDGE_RESIST { get; set; } 
        public string c73_DATE_CODE { get; set; } 
        public string c74_LOT_NUMBER { get; set; } 
        public string c75_STATION_NUM { get; set; } 
        public string c76_WORKERS { get; set; } 
        public string c77_WORKER2 { get; set; } 
        public string c78_DUPLICATE { get; set; } 
        public string c79_PIECE_NUM { get; set; } 
        public string c80_DELAY_FRACTION { get; set; } 
        public string c81_GENERAL { get; set; } 
        public string c82_GENERAL_1 { get; set; } 
        public string c83_BOX_PCS { get; set; } 
        public string c84_MANUAL { get; set; } 
        public string c85_PRODUCER { get; set; } 
        public string c86_DMX_PACKAGE_QTY { get; set; } 
        public string c87_DMX_DELAY_LEVEL { get; set; } 
        public string c88_DMX_DELAY_STEP { get; set; } 
        public string c89_DMX_NOMINAL_TIME_EA { get; set; } 
        public string c90_DMX_NOMINAL_TIME_NA { get; set; } 
        public string c91_DMX_WIRE_LENGTH { get; set; } 
        public string c92_DMX_ST_LENGTH { get; set; } 
        public string c93_DMX_DATE { get; set; } 
        public string c94_DMX_PACKAGE_NUM2 { get; set; } 
        public string c95_DMX_SEQ_NUM { get; set; } 
        public string c96_DMX_PRODUCER { get; set; } 
        public string c97_AUTO_LABEL { get; set; } 
        public string c98_TOTAL_PACKAGE_QTY { get; set; } 

    }
}
