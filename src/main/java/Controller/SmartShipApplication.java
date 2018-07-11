package Controller;

import Model.Constants;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;

import java.util.ArrayList;
import java.util.Map;

/**
 * Created by Ariel Peretz, Smartship company 2018
 */

//TODO if Products column hs DDP value - disregard it + log and report it
//TODO add log and report mechanism - for now print to console
//TODO change extra weight diff calculation to be fetched from price list instead of if else
//TODO check gui connection for app
//TODO create new WB or file for errors and unknown customers to be viewed manually

public class SmartShipApplication {

    public static void main(String[] args) {
        try {
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

            //copy invoice file not to work on original file
            UtilityMethods.copyFile(Constants.INVOICE_FILE_SOURCE_PATH, Constants.INVOICE_FILE_DES_PATH);

            //load invoice Workbook
            final Workbook invoiceWb = UtilityMethods.loadWb(Constants.INVOICE_FILE_DES_PATH);

            //sheet for copying sheet info to customer sheets
            Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHHET_NUM);
            //find customer-name column value
            final CellAddress cellCustomerName = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
            if (cellCustomerName == null) {
                throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
            }


            //Row for copying row info to customer rows
            final Row row = invoiceSheet.getRow(Constants.FIRST_ROW_NUM);

            //a mapping of customer names from to their file names
            final Map<String, String> customerNameToFileName = UtilityMethods.populateCustomerAndFileNames(invoiceWb);

            //map of customer File names to Workbooks
            final Map<String, Workbook> mapCustomerFileNameWb = UtilityMethods.createCustomerWbMap(row, customerNameToFileName);
            //create customer workbooks, file per customer name

            //find cell address of customer reference(id)
            final CellAddress customerIdCellAddress = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
            if (customerIdCellAddress == null) {
                throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
            }

            //copy rows to customer workbooks
            UtilityMethods.copyRowsToCustomersWb(invoiceWb, customerIdCellAddress, mapCustomerFileNameWb, customerNameToFileName);

            final Map<String, Integer> countryToRegionCodeMap = UtilityMethods.loadRegionToCountryMap(sheet, regionIndexCol, countryNameCol);


            //calc freight cell value from formula per customer
            UtilityMethods.calcFreightForAllCustomers(mapCustomerFileNameWb, countryToRegionCodeMap);//TODO DEBUGGING REACHED HERE

            //write files do disk
            UtilityMethods.saveAndCloseWbFiles(mapCustomerFileNameWb);

            System.out.println(" Finished Main ");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}