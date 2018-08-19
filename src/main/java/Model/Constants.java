package Model;

/**
 * Created by ARIELPE on 10/11/2017.
 */
public class Constants {

    public static final String SHEET_NAME = "Sheet1";
    public static final String INPUT_DIR = "inputdir";

    //file paths
    public static final String INVOICE_FILE_SOURCE_PATH = Constants.INPUT_DIR + "/DHL Invoices/March 3rd 2018.xlsx";
    public static final String REGION_TO_COUNTRY_FILE = Constants.INPUT_DIR + "/Region to country map/Regions to Country Mapping.xls";
    public static final String OUTPUT_DIR = "outputdir";
    public static final String INVOICE_FILE_DES_PATH = Constants.OUTPUT_DIR + "/invoiceFile workbook" + System.currentTimeMillis() + ".xls";
    public static final String OUT_LOG_FILE = Constants.OUTPUT_DIR + "/log files/log " + System.currentTimeMillis() + ".txt";
    public static final String EXCEPTION_LOG_FILE = "exceptions and logs " + System.currentTimeMillis() + ".txt";
    public static final String CUSTOMER_PRICE_LISTS = "customer price lists";

    //regular expressions
    public static final String REGEX_ONLY_NUMBERS = "[0-9]\\+";
    public static final String REGEX_FILTER_UNWANTED_CHARS = "[\\-\\+\\.\\^:,/]";

    //public static final String REGEX_ONLY_NUMBERS_A_TO_Z_LETTERS = "\\^[a-z0-9]\\+$/i";
    public static final String REGEX_ILLEGAL_CHARS = "[/.,!@\\\\#$>:;|<%^&?*()-]";
    public static final String REF_NUM_COLUMN_NAME = "Reference_no";
    public static final String BLANK = "";
    @Deprecated
    public static final String FREIGHT = "Frieght";

    public static final String OTHER = "Other";
    public static final String DDP = "DDP";
    public static final String DOX = "DOX";
    public static final String WPX = "WPX";
    public static final String COUNTRY_COL = "countryname";
    public static final String ZONE_NUM_COL = "ISRAEL ZONE";

    public static final String WEIGHT_COL = "Wght";
    public static final String DES_COL = "Destination";
    public static final String CUSTOMER_COLUMN_NAME = "Shipper_nm";
    public static final String FREIGHT_NEW_FORMAT_SHP = "_SHP";
    public static final String FUEL_NEW_FORMAT = "FF";
    public static final String PRODUCTS_COL = "Products";
    public static final String AWB_COL = "Awb";
    public static final String SHIP_DATE_COL = "Shp Date";
    public static final String ORIGIN_COL = "Origin";
    public static final String PERIOD_COL = "Period";
    public static final String REF_NUM_COL = "Reference_no";
    public static final String CNSGNEE_COL = "cnsgnee_nm";


    //customer price list files
    public static final String PL_FILE_ENDING = "price list.xlsx";
    public static final String XLSX_FILE_ENDING = ".xlsx";

    //Errors
    public static final String COLUMN_NOT_FOUND_ERROR = "Couldn't find column name";
    public static final String PRODUCTS_COLUMN_NOT_FOUND_ERROR = "Couldn't find Product column";
    public static final String CUSTOMER_COLUMN_NOT_FOUND_ERROR = "Couldn't find customer name column";
    public static final String FUEL_SURCHARGE_NOT_IN_RANGE = "FuelSurchargePercent not in the range of 0-1";
    public static final String COUNTRY_CODE_ERROR = "country code does not exist in table";
    public static final String WB_NOT_FOUND_ERROR = "Workbook does not exist in table";
    public static final String FREIGHT_NOT_FOUND_ERROR = "Frieght Cell not found";
    public static final String WEIGHT_NOT_FOUND_ERROR = "Weight Cell not found";
    public static final String DES_COL_NOT_FOUND_ERROR = "Destination Cell not found";
    //Log messages
    public static final String PRODUCT_IS_DDP = "Product Column has value of 'DDP' and therefor disregarded ";
    //DHL invoice file
    public static int FIRST_ROW_NUM = 0;
    public static int FIRST_SHEET_NUM = 0;
    public static int SECOND_SHEET_NUM = 1;

}
