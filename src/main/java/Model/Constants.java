package Model;

/**
 * Created by ARIELPE on 10/11/2017.
 */
public class Constants {

	public static final String SHEET_NAME = "Sheet1";

	//regular expressions
	public static final String REGEX_ONLY_NUMBERS = "[0-9]\\+";
	public static final String REGEX_FILTER_UNWANTED_CHARS = "[\\-\\+\\.\\^:,/]";
	public static final String REGEX_ONLY_NUMBERS_A_TO_Z_LETTERS = "\\^[a-z0-9]\\+$/i";
	public static final String REGEX_ILLEGAL_CHARS = "[/.,!@\\\\#$>:;|<%^&?*()-]";
	//public static final String REGEX_TMP = "[\\/:*?.,!@#$%^&-()\"<>|]";

	//DHL invoice file
	public static int FIRST_ROW_NUM = 0;
	public static int FIRST_SHHET_NUM = 0;
	public static final String CUSTOMER_COLUMN_NAME = "Shipper_nm";
	public static final String REF_NUM_COLUMN_NAME = "Reference_no";
	public static final String BLANK = "";
	public static final String FREIGHT = "Frieght";
	public static final String COUNTRY_COL = "countryname";
	public static final String ZONE_NUM_COL = "ISRAEL ZONE";
	public static final String WEIGHT_COL = "Wght";
	public static final String DES_COL = "Destination";

	//customer price list files
	public static final String PL_FILE_ENDING = "price list.xlsx";
	public static int ZONE_OFFSET = 3;
	public static double WEIGHT_MULTIPLIER = 0.5;


	//Errors
	public static final String COLUMN_NOT_FOUND_ERROR = "Couldn't find column name";
	public static final String FUEL_SURCHARGE_NOT_IN_RANGE = "FuelSurchargePercent not in the range of 0-1";
	public static final String COUNTRY_CODE_ERROR = "country code does not exist in table";

}
