package Model;

/**
 * Created by ARIELPE on 10/11/2017.
 */
public class Constants {

	public static String SHEET_NAME = "Sheet1";

	//regular expressions
	public static String REGEX_ONLY_NUMBERS = "[0-9]\\+"; //
	public static String REGEX_ONLY_NUMBERS_A_TO_Z_LETTERS = "\\^[a-z0-9]\\+$/i";
	public static String REGEX_ILLEGAL_CHARS = "[/.,!@\\\\#$>:;|<%^&?*()-]";
	//public static String REGEX_TMP = "[\\/:*?.,!@#$%^&-()\"<>|]";

	//DHL invoice file
	public static int FIRST_ROW_NUM = 0;
	public static String CUSTOMER_COLUMN_NAME = "Shipper_nm";
	public static String REF_NUM_COLUMN_NAME = "Reference_no";
	public static String BLANK = "";
	public static String FREIGHT = "Frieght";
	public static String COUNTRY_COL = "countryname";
	public static String ZONE_NUM_COL = "ISRAEL ZONE";
	public static String WEIGHT_COL = "Wght";
	public static String DES_COL = "Destination";

	//customer price list files
	public static String PL_FILE_ENDING = "price list.xlsx";
	public static int ZONE_OFFSET = 3;
	public static double WEIGHT_MULTIPLIER = 0.5;

	//customersNamesToIdsMapping.xls file
	public static int CUST_NAME_COL = 0;
	public static int CUST_ID_COL = 1;
	public static int CUST_TYPE_COL = 2;

	//Errors
	public static String COLUMN_NOT_FOUND_ERROR = "Couldn't find column name";
	public static String FUEL_SURCHARGE_NOT_IN_RANGE = "FuelSurchargePercent not in the range of 0-1";

}
