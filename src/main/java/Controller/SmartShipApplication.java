package Controller;

import Model.Constants;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.Map;

/**
 * Created by Ariel Peretz, Smartship 2018
 */

public class SmartShipApplication {


    /**
     * Main Application flow, running all activities one by one
     *
     * @param fuel
     * @return boolean when all activities finish
     */
    public static boolean applicationFlow(String fuel) {
        boolean success = true;

        try {
            //value of fuel surcharge
            final Double fuelSurcharge = Double.parseDouble(fuel);

            //changing standard output to write logs into a log file instead of console
            final FileOutputStream logFile = new FileOutputStream(Constants.EXCEPTION_LOG_FILE);
            final PrintStream printStream = new PrintStream(logFile);
            System.setOut(printStream);

            UtilityMethods utilityMethods = new UtilityMethods(printStream);

            //copy invoice file not to work on original file
            utilityMethods.copyFile(Constants.INVOICE_FILE_SOURCE_PATH, Constants.INVOICE_FILE_DES_PATH);

            //load invoice Workbook
            final Workbook invoiceWb = utilityMethods.loadWb(Constants.INVOICE_FILE_DES_PATH);

            //sheet for copying sheet info to customer sheets
            Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHEET_NUM);

            //Row for copying row info to customer rows
            final Row row = invoiceSheet.getRow(Constants.FIRST_ROW_NUM);

            //a mapping of customer names from to their file names
            final Map<String, String> customerNameToFileName = utilityMethods.populateCustomerAndFileNames(invoiceWb);

            //creates a map of customer File names to Workbooks, creates WB for each customer
            final Map<String, Workbook> mapCustomerFileNameWb = utilityMethods.createCustomerWbMap(row, customerNameToFileName);
            //create customer workbooks, file per customer name

            //find cell address of customer column name
            final CellAddress customerColumnCellAddress = utilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
            if (customerColumnCellAddress == null) {
                throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
            }

            //copy rows to customer workbooks
            utilityMethods.copyRowsToCustomersWb(invoiceWb, customerColumnCellAddress, mapCustomerFileNameWb, customerNameToFileName);

            final Workbook countryToRegionCodeWb = utilityMethods.loadWb(Constants.REGION_TO_COUNTRY_FILE);
            final Sheet sheet = countryToRegionCodeWb.getSheet(Constants.SHEET_NAME);
            //get region column
            final CellAddress cellRegionIndexCol = utilityMethods.findCellByName(sheet, Constants.ZONE_NUM_COL);
            if (cellRegionIndexCol == null) {
                throw new IllegalArgumentException(Constants.COLUMN_NOT_FOUND_ERROR);
            }
            final int regionIndexCol = cellRegionIndexCol.getColumn();

            //get country column
            final CellAddress cellCountryNameCol = utilityMethods.findCellByName(sheet, Constants.COUNTRY_COL);
            if (cellCountryNameCol == null) {
                throw new IllegalArgumentException(Constants.COLUMN_NOT_FOUND_ERROR);
            }
            final int countryNameCol = cellCountryNameCol.getColumn();

            final Map<String, Integer> countryToRegionCodeMap = utilityMethods.loadRegionToCountryMap(sheet, regionIndexCol, countryNameCol);

            //calculate shipment cost(freight) for customers
            utilityMethods.calcFreightForAllCustomers(mapCustomerFileNameWb, countryToRegionCodeMap, fuelSurcharge);

            //page setup
            utilityMethods.pagePrintSetup(mapCustomerFileNameWb);

            //write files do disk
            utilityMethods.saveAndCloseWbFiles(mapCustomerFileNameWb);


            logFile.close();
            System.out.println(" Finished SmartShipApplication main ");
        } catch (Exception e) {
            success = false;
            e.printStackTrace();
        }
        return success;
    }
}