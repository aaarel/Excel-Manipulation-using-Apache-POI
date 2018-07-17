package Controller;

import Model.Constants;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import java.util.Map;

/**
 * Created by Ariel Peretz, Smartship 2018
 */

public class SmartShipApplication {

    public static void main(String[] args) {

        //copy invoice file not to work on original file
        UtilityMethods.copyFile(Constants.INVOICE_FILE_SOURCE_PATH, Constants.INVOICE_FILE_DES_PATH);

        //load invoice Workbook
        final Workbook invoiceWb = UtilityMethods.loadWb(Constants.INVOICE_FILE_DES_PATH);

        //sheet for copying sheet info to customer sheets
        Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHEET_NUM);

        //Row for copying row info to customer rows
        final Row row = invoiceSheet.getRow(Constants.FIRST_ROW_NUM);

        //a mapping of customer names from to their file names
        final Map<String, String> customerNameToFileName = UtilityMethods.populateCustomerAndFileNames(invoiceWb);

        //creates a map of customer File names to Workbooks, creates WB for each customer
        final Map<String, Workbook> mapCustomerFileNameWb = UtilityMethods.createCustomerWbMap(row, customerNameToFileName);
        //create customer workbooks, file per customer name

        //find cell address of customer column name
        final CellAddress customerColumnCellAddress = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
        if (customerColumnCellAddress == null) {
            throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
        }

        //copy rows to customer workbooks
        UtilityMethods.copyRowsToCustomersWb(invoiceWb, customerColumnCellAddress, mapCustomerFileNameWb, customerNameToFileName);

        final Workbook countryToRegionCodeWb = UtilityMethods.loadWb(Constants.REGION_TO_COUNTRY_FILE);
        final Sheet sheet = countryToRegionCodeWb.getSheet(Constants.SHEET_NAME);
        //get region column
        final CellAddress cellRegionIndexCol = UtilityMethods.findCellByName(sheet, Constants.ZONE_NUM_COL);
        if (cellRegionIndexCol == null) {
            throw new IllegalArgumentException(Constants.COLUMN_NOT_FOUND_ERROR);
        }
        final int regionIndexCol = cellRegionIndexCol.getColumn();

        //get country column
        final CellAddress cellCountryNameCol = UtilityMethods.findCellByName(sheet, Constants.COUNTRY_COL);
        if (cellCountryNameCol == null) {
            throw new IllegalArgumentException(Constants.COLUMN_NOT_FOUND_ERROR);
        }
        final int countryNameCol = cellCountryNameCol.getColumn();

        final Map<String, Integer> countryToRegionCodeMap = UtilityMethods.loadRegionToCountryMap(sheet, regionIndexCol, countryNameCol);

        //calculate shipment cost(freight) for customers
        UtilityMethods.calcFreightForAllCustomers(mapCustomerFileNameWb, countryToRegionCodeMap);

        //insert img        // add Smartship logo
        //UtilityMethods.insertLogo(mapCustomerFileNameWb);

        //page setup
        UtilityMethods.pagePrintSetup(mapCustomerFileNameWb);

        // convert to PDF

        //write files do disk
        UtilityMethods.saveAndCloseWbFiles(mapCustomerFileNameWb);

        //load files and delete cols save files
        AsposeCellsUtilities.deleteBlankAndAutoFitColumns(mapCustomerFileNameWb);

        //load files remove new sheet save and close wb files
        //UtilityMethods.deleteUnusedSheetsInWb(mapCustomerFileNameWb);


        //export log/error file

        System.out.println(" Finished Main ");


    }
}