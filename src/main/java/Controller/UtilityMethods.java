package Controller;

import Model.Constants;
import Model.Customer;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class UtilityMethods {

    //copies rows from invoice WB sheet to des customer WB sheet //TODO slow - make faster use threads in parallel ?
    public static void copyRowsToCustomersWb(Workbook wb, CellAddress customerIdCellAddress, Map<String, Workbook> mapCustomerToWb, Map<String, String> customerNameToFileName) {
        //iterate over all sheets in workbook
        for (int sheetCount = 0; sheetCount < wb.getNumberOfSheets(); sheetCount++) {
            //add rows from main wb to to customers WBs
            final Sheet originalSheet = wb.getSheetAt(sheetCount);
            Row originalRow;
            Workbook customerWb;
            Sheet customerSheet;
            Row customerRow;
            String customerName;
            String customerFileName;
            final DataFormatter dataFormatter = new DataFormatter();
            //iterate over all cells in workbook
            //TODO add here check for DDP in products cell
            for (int i = 1; i <= originalSheet.getLastRowNum(); i++) {
                originalRow = originalSheet.getRow(i);
                customerName = dataFormatter.formatCellValue(originalRow.getCell(customerIdCellAddress.getColumn()));
                customerFileName = customerNameToFileName.get(customerName);
                if (customerFileName == null) {
                    throw new NullPointerException(Constants.WB_NOT_FOUND_ERROR);
                }
                //get current customer wb, sheet
                customerWb = mapCustomerToWb.get(customerFileName);
                if (customerWb == null) {
                    throw new NullPointerException(Constants.WB_NOT_FOUND_ERROR);
                }
                customerSheet = customerWb.getSheet(Constants.SHEET_NAME);

                //create Row in new sheet
                int customerRowIndex = customerSheet.getLastRowNum() + 1;
                customerRow = customerSheet.createRow(customerRowIndex);

                //copy Rows(cell by cell) to customer sheet
                for (int j = 0; j < originalRow.getLastCellNum(); j++) {    //TODO better to use {while cellIterator.hasNext()} ?
                    Cell cell = customerRow.createCell(j);
                    Cell originalCell = originalRow.getCell(j);

                    //Copy style from old cell and apply to new cell
                    //HSSFCellStyle newCellStyle = workbook.createCellStyle();
                    //newCellStyle.cloneStyleFrom(originalCell.getCellStyle()); //TODO fix styling of cells later
                    //cell.setCellStyle(newCellStyle);

                    // If there is a cell comment, copy
                    if (originalCell.getCellComment() != null) {
                        cell.setCellComment(originalCell.getCellComment());
                    }
                    // If there is a cell hyperlink, copy
                    if (originalCell.getHyperlink() != null) {
                        cell.setHyperlink(originalCell.getHyperlink());
                    }

                    // Set the cell data type
                    cell.setCellType(originalCell.getCellType());
                    // Set the cell data value
                    switch (originalCell.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            cell.setCellValue(originalCell.getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            cell.setCellValue(originalCell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            cell.setCellErrorValue(originalCell.getErrorCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            cell.setCellFormula(originalCell.getCellFormula());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            cell.setCellValue(originalCell.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_STRING:
                            cell.setCellValue(originalCell.getRichStringCellValue());
                            break;
                    }
                }
                System.out.println("copied originalRow with index: " + customerRowIndex);
            }
            System.out.println("Finished sheet " + originalSheet.getSheetName());
        }
        System.out.println("Finished iterating workbook ");
        System.out.println("******************************");

    }

    //saves and closes open WB files
    public static void saveAndCloseWbFiles(Map<String, Workbook> mapCustomerFileNameToWb) throws IOException {
        //save and close workbooks and files
        for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
            //close and save
            FileOutputStream fileOutCustomer = new FileOutputStream("outputdir/customers/" + customerFileName + ".xls");
            mapCustomerFileNameToWb.get(customerFileName).write(fileOutCustomer);
            System.out.println(" Wrote workbook file to disk: " + customerFileName);
            fileOutCustomer.close();
            System.out.println(" Closed file: " + customerFileName + ".xls");
        }
        System.out.println("******************************");
    }

    //returns a map of Customer to workbook, creates a new WB file per customer including wb headlines
    public static Map<String, Workbook> createCustomerWbMap(Row firstRow, Map<String, String> customerNameToFileName) {
        final Map<String, Workbook> mapCustomerFileNameWb = new HashMap<String, Workbook>();
        for (String customer : customerNameToFileName.values()) {
            mapCustomerFileNameWb.put(customer, createCustomerWb(firstRow));
            System.out.println("created customer wb with name: " + customer);
        }
        System.out.println("******************************");
        return mapCustomerFileNameWb;
    }

    //creates a wb with a row used as headline
    public static Workbook createCustomerWb(Row firstRow) {
        final Workbook wb = new XSSFWorkbook();
        final Sheet wbSheet = wb.createSheet(Constants.SHEET_NAME);
        final Row wbRow = wbSheet.createRow(Constants.FIRST_ROW_NUM);
        //fill first row with headlines
        for (int cellCount = 0; cellCount < firstRow.getLastCellNum(); cellCount++) {
            wbRow.createCell(cellCount).setCellValue(firstRow.getCell(cellCount).getStringCellValue());
        }
        return wb;
    }

    //prints all cell data from a given sheet
    private static void printDataFromSheet(Sheet hssfSheet) {
        //iterate over sheet rows
        Iterator<Row> rowIterator = hssfSheet.iterator();
        //loop through sheet
        while (rowIterator.hasNext()) {
            String name = "";
            String shortCode = "";

            //Get the row object
            Row row = rowIterator.next();

            //Every row has columns, get the column iterator and iterate over them
            Iterator<Cell> cellIterator = row.cellIterator();

            //loop through cells
            while (cellIterator.hasNext()) {
                //Get the Cell object
                Cell cell = cellIterator.next();
                //check the cell type and process accordingly
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        if (shortCode.equalsIgnoreCase("")) {
                            shortCode = cell.getStringCellValue().trim();
                        } else if (name.equalsIgnoreCase("")) {
                            //2nd column
                            name = cell.getStringCellValue().trim();
                        } else {
                            //random data, leave it
                            System.out.println("Random data::" + cell.getStringCellValue());
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.println("Random data::" + cell.getNumericCellValue());
                }
            } //end of cell iterator

        } //end of rows iterator
    }

    //finds cell address by name
    public static CellAddress findCellByName(Sheet hssfSheet, String customerColumnName) {
        //row iterator
        Iterator<Row> rowIterator = hssfSheet.iterator();
        Iterator<Cell> cellIterator;
        Row row;
        Cell cell;
        DataFormatter dataFormatter = new DataFormatter();
        while (rowIterator.hasNext()) {
            //get line
            row = rowIterator.next();
            cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                //get cell
                cell = cellIterator.next();
                //check if cell string is customer name column
                if (dataFormatter.formatCellValue(cell).equals(customerColumnName)) {
                    System.out.println("Found Cell, Column: " + cell.getColumnIndex() + " Row: " + cell.getRowIndex());
                    CellAddress cellAddress = cell.getAddress();
                    System.out.println("cellAddress : " + cellAddress.toString());
                    return cellAddress;
                }
            }
        }
        return null;
    }

    //gets customer IDs from one sheet and changes cell values to correct names
    public static HashSet<String> getCustomerIdsFromSheet(Sheet sheet, int columnIndex) {
        final HashSet<String> customerIdsSet = new HashSet<String>();
        Iterator<Row> rowIterator = sheet.iterator();
        Iterator<Cell> cellIterator;
        DataFormatter dataFormatter = new DataFormatter();
        Cell cell;
        Row row;
        String cellVal;
        //iterate over rows
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            cellIterator = row.cellIterator();
            //iterate over cells
            while (cellIterator.hasNext()) {
                cell = cellIterator.next();
                //checking if in proper column and in values section
                if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
                    //validates and cleans cell value if, customer ID is allowd : [0-9]
                    cellVal = dataFormatter.formatCellValue(cell);
                    if (!cellVal.equals(cellVal.replaceAll(Constants.REGEX_ONLY_NUMBERS, Constants.BLANK))) {
                        System.out.println("old Cell Value: " + cellVal + " changed to new cell value: " + cellVal.replaceAll(Constants.REGEX_ONLY_NUMBERS, Constants.BLANK));
                        cell.setCellValue(cellVal.replaceAll(Constants.REGEX_ONLY_NUMBERS, Constants.BLANK)); //TODO correct place to change cell values ?
                    }
                    customerIdsSet.add(dataFormatter.formatCellValue(cell));
                    break;
                }
            }
        }
        return customerIdsSet;
    }

    //gets customer names from one sheet, "clean" names from unwanted characters
    public static Map<String, String> getCustomerNamesAndFileNamesFromSheet(Sheet hssfSheet, int columnIndex) {
        Map<String, String> customerNameToFile = new HashMap<String, String>();
        Iterator<Row> rowIterator = hssfSheet.iterator();
        Iterator<Cell> cellIterator;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                //checking if in proper column and in values section
                if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
                    String customerNameAsInSheet = cell.getStringCellValue();
                    //fileName= customer name in upper case -minus illegal characters
                    String customerFileName = customerNameAsInSheet.replaceAll(Constants.REGEX_FILTER_UNWANTED_CHARS, " ").toUpperCase();
                    if (!customerNameToFile.containsKey(customerNameAsInSheet)) {
                        customerNameToFile.put(customerNameAsInSheet, customerFileName);
                        System.out.println("Added customer: " + customerNameAsInSheet + " with file name: " + customerFileName);
                    }

                    break;
                }
            }
        }
        return customerNameToFile;
    }

    //copies file
    public static void copyFile(String srcPath, String desPath) {
        try {
            FileUtils.copyFile(new File(srcPath), new File(desPath));
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("copied file: " + srcPath + " successfully");
    }

    //loads workbook
    public static Workbook loadWb(String path) {
        //read file contents
        File file = new File(path);
        try {
            FileInputStream fIP = new FileInputStream(file);
            //Get the Workbook instance for XLS file
            Workbook workbook = WorkbookFactory.create(fIP);
            if (file.isFile() && file.exists()) {
                System.out.println(path + " Open wb successfully ");
            } else {
                System.out.println(" Error Opening " + path);
            }
            return workbook;
        } catch (Exception e) {
            e.printStackTrace();
        }
        throw new IllegalStateException();
    }

    //read customer ids file, returns ID to Name map
    public static Map<String, Integer> loadRegionToCountryMap(Sheet sheet, int idColumn, int countryNameCol) {
        Row row;
        final Map<String, Integer> regionToCountryMap = new HashMap<String, Integer>();
        final DataFormatter dataFormatter = new DataFormatter();

        //iterate over all rows and cells in workbook
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            String countryName = dataFormatter.formatCellValue(row.getCell(countryNameCol));
            String regionId = dataFormatter.formatCellValue(row.getCell(idColumn));
            regionToCountryMap.put(countryName, Integer.parseInt(regionId));
        }
        return regionToCountryMap;
    }

    //calculates freight for all customers in map
    public static void calcFreightForAllCustomers(Map<String, Workbook> mapCustomerFileNameWb, Map<String, Integer> regionsMap) {
        for (String customer : mapCustomerFileNameWb.keySet()) {
            //load customer price list
            final Workbook customerPriceListWb = loadWb(Constants.INPUT_DIR + "/" + Constants.CUSTOMER_PRICE_LISTS + "/" + customer + Constants.XLSX_FILE_ENDING);
            calcCustomerFreight(customer, mapCustomerFileNameWb.get(customer), customerPriceListWb, regionsMap);// TODO check not null
        }
        System.out.println("******************************");
    }

    /**
     * calculates freight cell value according to formula
     * formula: CHARGE = [(customer price per region) X weight]  + [(fuel surcharge% * customer price per region)]
     * //TODO insurance value add with
     * //TODO fuel surecharge is dynamic - how to get externally every time
     */
    public static void calcCustomerFreight(String customer, Workbook wb, Workbook customerPriceListWb, Map<String, Integer> regionsMap) {
        final Sheet sheet = wb.getSheet(Constants.SHEET_NAME);
        //init cells
        int weightCol = -1;
        int destinationCol = -1;
        int freightCol = -1;
        int productsCol = -1;

        try {
            System.out.println("Started calcCustomerFreight on : " + customer);

            //get Frieght cell value
            CellAddress cellFrieght = findCellByName(sheet, Constants.FREIGHT);
            if (cellFrieght == null) {
                cellFrieght = findCellByName(sheet, Constants.FREIGHT_NEW_FORMAT);
            }
            if (cellFrieght == null) {
                throw new IllegalArgumentException(Constants.FREIGHT_NOT_FOUND_ERROR);
            }
            freightCol = cellFrieght.getColumn();

            //get Weight cell value
            CellAddress cellWeight = findCellByName(sheet, Constants.WEIGHT_COL);
            if (cellWeight == null) {
                throw new IllegalArgumentException(Constants.WEIGHT_NOT_FOUND_ERROR);
            }
            weightCol = cellWeight.getColumn();

            //get Destination cell value
            CellAddress cellDes = findCellByName(sheet, Constants.DES_COL);
            if (cellDes == null) {
                throw new IllegalArgumentException(Constants.DES_COL_NOT_FOUND_ERROR);
            }
            destinationCol = cellDes.getColumn();

            CellAddress cellProduct = findCellByName(sheet, Constants.PRODUCTS_COL);
            if (cellDes == null) {
                throw new IllegalArgumentException(Constants.DES_COL_NOT_FOUND_ERROR);
            }
            productsCol = cellProduct.getColumn();

            //TODO insurance col if exists
            //TODO Rishomon log
        } catch (Exception e) {
            e.printStackTrace();
        }
        //TODO get fuelSC %
        // via web {http://www.dhl.co.il/en/express/shipping/shipping_advice/express_fuel_surcharge_eu.html}
        final double fuelScp = getFuelSurchargePercent();
        if (fuelScp < 0 || fuelScp > 1) {
            throw new IllegalArgumentException(Constants.FUEL_SURCHARGE_NOT_IN_RANGE);
        }

        DataFormatter dataFormatter = new DataFormatter();
        Row row;
        Cell cell;
        double weight;
        String country;
        int zone;
        double pricePerWeightAndZone;
        double totalPrice;
        //iterate over all rows in customer workbook ( 1 row = 1 shipment)
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            //get current row
            row = sheet.getRow(i);

            cell = row.getCell(productsCol);
            if (cell.getStringCellValue() == Constants.DDP) {
                System.out.println(Constants.PRODUCT_IS_DDP + " customer WB: " + customer + " row: " + row.getRowNum());
                continue;   //DDP is disregarded, jump to next row
            }

            //get weight value
            cell = row.getCell(weightCol);
            weight = Double.parseDouble(dataFormatter.formatCellValue(cell));
            //get shipment des country
            cell = row.getCell(destinationCol);
            country = dataFormatter.formatCellValue(cell);

            //get zone for country
            if (regionsMap.containsKey(country)) {
                zone = regionsMap.get(country);
            } else if (regionsMap.containsKey(country.toUpperCase())) {
                zone = regionsMap.get(country.toUpperCase());
            } else if (regionsMap.containsKey(country.toLowerCase())) {
                zone = regionsMap.get(country.toLowerCase());
            } else {
                System.out.println("Country not found: "+ country);
                throw new IllegalArgumentException(Constants.COUNTRY_CODE_ERROR);
            }

            //calc price according to weight and zone
            pricePerWeightAndZone = getCustomerPrice(customerPriceListWb.getSheet(Constants.SHEET_NAME), weight, zone);
            totalPrice = ((1 + fuelScp) * pricePerWeightAndZone);                    //fuelScp range:[0 - 1]

            //write updated price value to cell
            cell = row.getCell(freightCol);
            cell.setCellValue(totalPrice);
            System.out.println("Updated cell : " + cell.getAddress() + " with a new FRIEGHT value of: " + totalPrice);
        }
        System.out.println("******************************");
    }

    /**
     * gets the price per zone for a given customer
     * //find closest base weight and get price example: weight is 11.5 , base =11K  D5-d53
     * Zone Cells: E4-K4, values: [1-7], ZONE_OFFSET = 3
     */
    public static double getCustomerPrice(Sheet priceListSheet, double weight, int zone) {
        //rows and columns are Zero based
        Row row, nextRow;
        Cell cell;
        double baseWeight = 0;
        double basePrice = 0;
        double diff;
        double additionalPrice = 0;
        double nextStepPrice = 0;
        final int zoneCol = zone; //+ Constants.ZONE_OFFSET;
        //row range: 2-28
        final int startRow = 2;
        final int endRow = 28;
        //weight column: 0
        final int baseWeightCol = 0;

        //TODO use (binary) search for sorted values and search more efficient by log(n)
        //iterate over rows

        int i;
        //go over rows in price list

        for (i = startRow; i <= endRow; i++) {
            row = priceListSheet.getRow(i);
            cell = row.getCell(baseWeightCol);
            baseWeight = cell.getNumericCellValue();

            if (weight == baseWeight) {
                basePrice = row.getCell(zoneCol).getNumericCellValue();
                System.out.println("Found exact weight: " + weight + " Price is: " + basePrice);
                break;
            }

            //not to overflow
            else if ((weight > baseWeight) && (i + 1 <= endRow)) {
                baseWeight = cell.getNumericCellValue();
                nextRow = priceListSheet.getRow(i + 1);
                //find closest weight value
                if ((weight >= cell.getNumericCellValue()) && (weight < priceListSheet.getRow(i + 1).getCell(baseWeightCol).getNumericCellValue())) {
                    basePrice = row.getCell(zoneCol).getNumericCellValue(); //price for base wight
                    nextStepPrice = nextRow.getCell(zoneCol).getNumericCellValue(); //price for next step
                    additionalPrice = (weight - baseWeight) * (basePrice + nextStepPrice) / 2; //additional price =  remaining weight * avg price
                    basePrice = basePrice + additionalPrice;
                    System.out.println("Found  weight: " + weight + " Price is: " + basePrice);
                    break;
                }
            }
            //no price in price list for current weight //TODO all edge cases?
            else {
                System.out.println("weight: " + weight + " not found in price list table ");
                throw new IllegalArgumentException(Constants.WEIGHT_NOT_FOUND_ERROR);
            }
        }
        return basePrice;
    }

    //returns fuel surcharge
    public static double getFuelSurchargePercent() {
        //TODO
        //get getFuelSurchargePercent from somewhere
        return 0;
    }

    //TODO - DEPRECATED METHODS
    //prints workbook info (sheet names and numbers)
    public static void printWorkbookInfo(Workbook workbook) {
        //Headlines
        System.out.println(" Workbook Data: ");
        System.out.println(" ------------------------------------ ");

        //workbook
        //String workbookName = workbook.getName();
        //System.out.println(" Workbook name is: " + workbookName);

        //sheets in workbook
        int numOfSheets = workbook.getNumberOfSheets();
        System.out.println("in workbook: " + "there are: " + numOfSheets + " sheets ");

        //loop through sheets
        for (int i = 0; i < numOfSheets; i++) {
            System.out.println(" Sheet name: " + workbook.getSheetName(i) + " , Sheet number: " + i);
            Sheet sheet = workbook.getSheetAt(i);
        }
    }

    /**
     * copy customer rows from main sheet to customer sheets
     * iterate over rows ,for each row:
     * get cust name
     * if cust xls exist 	-> copy row to cust xls
     * else 				-> create new xls, new sheet, copy row to cust xls
     * copy row needs to check that the destination row does not exist before creating it
     * save & close all files
     *
     * @param mainSheet
     * @param customerNameColumn
     * @param customersSet
     */
    public static void copyCustomerRowsToCustomerSheet(Sheet mainSheet, int customerNameColumn, Set<String> customersSet) {
        //iterate over sheet rows
        Iterator<Row> rowIterator = mainSheet.iterator();
        //loop through rows in sheet
        while (rowIterator.hasNext()) {

            //Get the row object
            Row row = rowIterator.next();

            //if first row jump to next iteration
            int rowNum = row.getRowNum();
            if (rowNum == 0) {
                continue;
            }

            //Every row has columns, get the column iterator and iterate over them
            Iterator<Cell> cellIterator = row.cellIterator();

            //check customer
            String customerName = row.getCell(customerNameColumn).getStringCellValue();

            String pathToFile = "outputdir/customers/workbook " + customerName;
            try {
                Workbook customerWorkbook = loadWb(pathToFile);
                Sheet customerSheet = customerWorkbook.getSheet(customerName);
                Row customerRow = customerSheet.createRow(rowNum);
                //copy rows:
                //copyRows(row, customerRow); // TODO come back here later
                String name = "";
                String shortCode = "";
                //loop through cells
                while (cellIterator.hasNext()) {
                    //Get the Cell object
                    Cell cell = cellIterator.next();
                    //check the cell type and process accordingly
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:
                            if (shortCode.equalsIgnoreCase("")) {
                                shortCode = cell.getStringCellValue().trim();
                            } else if (name.equalsIgnoreCase("")) {
                                //2nd column
                                name = cell.getStringCellValue().trim();
                            } else {
                                //random data, leave it
                                System.out.println("Random data::" + cell.getStringCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println("Random data::" + cell.getNumericCellValue());
                    }
                } //end of cell iterator

            } catch (Exception e) {
                e.printStackTrace();
            }


        } //end of rows iterator
    }

    //copies ROW
    private static void copyRow(Workbook workbook, Sheet worksheet, Sheet resultSheet, int sourceRowNum, int destinationRowNum) {
        Row newRow = resultSheet.getRow(destinationRowNum);

        Row sourceRow = worksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            resultSheet.shiftRows(destinationRowNum, resultSheet.getLastRowNum(), 1);
        } else {
            newRow = resultSheet.createRow(destinationRowNum);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }
    }

    //returns list of columns strings
    public static List<String> getStringListOfColumns(Sheet hssfSheet) {
        //list of columns
        List<String> columnList = new ArrayList<String>();
        //row iterator
        Iterator<Row> rowIterator = hssfSheet.iterator();
        Iterator<Cell> cellIterator;
        Row row;
        Cell cell;
        while (rowIterator.hasNext()) {
            //get line
            row = rowIterator.next();
            while (row.cellIterator().hasNext()) {
                //get cell
                cellIterator = row.cellIterator();
                cell = cellIterator.next();
                //add to list
                columnList.add(cell.getStringCellValue());
                //get cell address
                cell.getAddress();
                //get column index
                cell.getColumnIndex();
            }

        }
        //
        return columnList;
    }

    //prints data from workbook
    public static void printDataFromWorkbook(Workbook workbook) throws Exception {

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet hssfSheet = workbook.getSheetAt(i);
            System.out.println(" Sheet name: " + hssfSheet.getSheetName() + " , Sheet number: " + i);
            printDataFromSheet(hssfSheet);
        }
    }

    //finds customer name column
    public static CellRangeAddress findCellRangeAddress(Sheet hssfSheet, CellAddress customerNameCellAddress) throws ClassNotFoundException {
        CellRangeAddress cellRangeAddress;
        int firstRow, lastRow, firstCol, lastCol;
        boolean firstRowFlag, lastRowFlag, firstColFlag, lastColFlag;

        //row iterator
        Iterator<Row> rowIterator = hssfSheet.iterator();
        Iterator<Cell> cellIterator;
        Row row;
        Cell cell;
        while (rowIterator.hasNext()) {
            //get line
            row = rowIterator.next();

            cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                //get cell
                cell = cellIterator.next();
                //check if cell string is customer name column
                if (cell.getStringCellValue().equals(Constants.CUSTOMER_COLUMN_NAME)) {
                    System.out.println("Found Cell, Column: " + cell.getColumnIndex() + " Row: " + cell.getRowIndex());
                    CellAddress cellAddress = cell.getAddress();
                    System.out.println("cellAddress : " + cellAddress.toString());
                    return null;
                }
            }
        }
        throw new ClassNotFoundException(Constants.COLUMN_NOT_FOUND_ERROR);
    }

    //creates workbook
    public static Workbook createWorkbook() throws Exception {
        //Create Blank Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create file system using specific name
        FileOutputStream out = new FileOutputStream(new File("outputdir/createWorkbook.xlsx"));
        //write operation Workbook using file out object
        workbook.write(out);
        out.close();
        System.out.println("createWorkbook.xlsx written successfully");
        return workbook;
    }

    //gets customer names from workbook
    public static HashSet<String> getCustomerNamesFromWorkbook(Workbook workbook, CellAddress cellAddress) {
        List<String> customerList = new ArrayList();
        int numOfSheets = workbook.getNumberOfSheets();
        Sheet hssfSheet;
        int columnIndex = cellAddress.getColumn();
        for (int i = 0; i < numOfSheets; i++) {
            hssfSheet = (Sheet) workbook.getSheetAt(i);
            //customerList.addAll(getCustomerNamesAndFileNamesFromSheet(hssfSheet, columnIndex));
        }
        return new HashSet<String>(customerList);
    }

    //clean illegal chars
    public static void removeIllegalCharactersFromColumnInSheet(Sheet sheet, int columnIndex) {
        Iterator<Row> rowIterator = sheet.iterator();
        Iterator<Cell> cellIterator;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                //checking if in proper column and in values section
                if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
                    //remove un-needed characters from name
                    cell.setCellValue(cell.getStringCellValue().replaceAll("[\\-\\+\\.\\^:,/]", ""));
                    break;
                }
            }
        }
    }

    //gets customer IDs from workbook
    public static HashSet<Customer> getCustomerIdsFromWorkbook(Workbook workbook, CellAddress customerIdCellAddress, Map<String, String> customerIdToNameMap) {
        final int numOfSheets = workbook.getNumberOfSheets();
        final int columnIndex = customerIdCellAddress.getColumn();
        final HashSet<String> customerIdsSet = new HashSet<String>();
        final HashSet<Customer> customerSet = new HashSet<Customer>();
        Sheet sheet;

        //iterate over sheets on wb, for each sheet clean, validate and add customer IDs to List
        for (int i = 0; i < numOfSheets; i++) {
            sheet = workbook.getSheetAt(i);
            customerIdsSet.addAll(getCustomerIdsFromSheet(sheet, columnIndex));
        }

        //create set of customers from Set of customer IDs(enrich)
        for (String id : customerIdsSet) {
            customerSet.add(new Customer(id, customerIdToNameMap.get(id)));
        }
        return customerSet;
    }

    //read customer ids file, returns ID to Name map
    public static Map<String, String> loadCustomerIdToNameMap(Sheet sheet, int idColumn, int nameColumn) {
        Row row;
        String customerId;
        String customerName;
        Map<String, String> idToNameMap = new HashMap<String, String>();
        DataFormatter dataFormatter = new DataFormatter();

        //iterate over all cells in workbook
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            //validate format, change names
            customerId = dataFormatter.formatCellValue(row.getCell(idColumn)).replaceAll(Constants.REGEX_ONLY_NUMBERS, "");
            //TODO improve reg-ex to include only a-z chars and numbers
            customerName = dataFormatter.formatCellValue(row.getCell(nameColumn)).replaceAll(Constants.REGEX_ILLEGAL_CHARS, "");
            if (!idToNameMap.containsKey(customerId)) {
                idToNameMap.put(customerId, customerName);
            }
        }
        return idToNameMap;
    }

    //creates(copies an example) a price list file per customer in the input workbook,
    //example call:         UtilityMethods.createPriceListFiles(new ArrayList<String>(customerNameToFileName.values()));
    public static void createPriceListFiles(List<String> names) {
        //create files price list per customer
        File folder = new File("inputdir/customer price lists/");
        List<String> list = Arrays.asList(folder.list());
        for (String s : names) {
            if (!list.contains(s + " " + Constants.PL_FILE_ENDING)) {
                copyFile("inputdir/customer price lists/EXAMPLE CUSTOMER PRICE LIST.xlsx", "inputdir/customer price lists/" + s + Constants.XLSX_FILE_ENDING);
                System.out.println("created file for: " + s);
            } else {
                System.out.println("file: " + s + " Already exists");
            }
        }
    }

    //populates map of Customer names from workbook to customer File Names (that cannot have illegal characters )
    public static Map<String, String> populateCustomerAndFileNames(Workbook invoiceWb) {
        Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHHET_NUM);
        final int customerColName = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME).getColumn();
        final Map<String, String> customerNameToFileName = new HashMap<String, String>();
        //get customer names from invoice wb and map to file names
        for (int i = 0; i < invoiceWb.getNumberOfSheets(); i++) {
            invoiceSheet = invoiceWb.getSheetAt(i);
            customerNameToFileName.putAll(UtilityMethods.getCustomerNamesAndFileNamesFromSheet(invoiceSheet, customerColName));
        }
        return customerNameToFileName;
    }

    //log and report method to check if each customer in the map has a respective price list file
    public static void printPriceListFilesInfo(Map<String, String> customerNameToFileName) {
        //get names of customer price list folder
        final File folder = new File(Constants.INPUT_DIR + "/" + Constants.CUSTOMER_PRICE_LISTS + "/");
        final File[] files = folder.listFiles();
        final Set<String> priceListFileNames = new HashSet<String>();
        System.out.println("******************************");
        System.out.println("priceListFiles... ");
        for (File file : files) {
            priceListFileNames.add(file.getName());
            System.out.println("price List File: " + file.getName());
        }
        System.out.println("******************************");

        //check if there are price lists per customer and log
        System.out.println("******************************");
        System.out.println("customer names not in price list ");
        for (String fileName : customerNameToFileName.values()) {
            if (!priceListFileNames.contains(fileName + Constants.XLSX_FILE_ENDING)) {
                System.out.println(fileName + " is not in priceListFileNames");
            }
        }
        System.out.println("******************************");
    }

}
