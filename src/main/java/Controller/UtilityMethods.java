package Controller;

import Model.Constants;
import Model.Customer;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.*;

public class UtilityMethods {

    //sets page print setup and closes open WB files
    public static void pagePrintSetup(Map<String, Workbook> mapCustomerFileNameToWb) {
        Workbook wb;
        Sheet sheet;
        try {
            //load files
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                //FileOutputStream fileOutCustomer = new FileOutputStream(Constants.OUTPUT_DIR + "/customers/" + customerFileName + ".xls");
                wb = mapCustomerFileNameToWb.get(customerFileName);
                sheet = wb.getSheet(Constants.SHEET_NAME); //has a problem here the new sheet has no index
                sheet.setFitToPage(true);
                sheet.getPrintSetup().setLandscape(true);
                sheet.getPrintSetup().setFitHeight((short) 1);
                sheet.getPrintSetup().setFitWidth((short) 1);
                //wb.write(fileOutCustomer);
                System.out.println(" Wrote workbook file to disk: " + customerFileName);
                //fileOutCustomer.close();
                System.out.println(" Closed file: " + customerFileName + ".xls");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("******************************");
    }


    //returns a list of approved column ids, approved = viewable by customer
    public static Set<Integer> approvedColIdList(Sheet sheet) {
        Set<Integer> approvedColumnList = new HashSet<Integer>();
        String val;
        Row row = sheet.getRow(0);
        Cell cell;
        DataFormatter dataFormatter = new DataFormatter();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            cell = cellIterator.next();
            val = dataFormatter.formatCellValue(cell);
            if (val.equals(Constants.AWB_COL) || val.equals(Constants.SHIP_DATE_COL) || val.equals(Constants.PRODUCTS_COL) ||
                    val.equals(Constants.ORIGIN_COL) || val.equals(Constants.DES_COL) || val.equals(Constants.PERIOD_COL) ||
                    val.equals(Constants.CUSTOMER_COLUMN_NAME) || val.equals(Constants.REF_NUM_COLUMN_NAME) || val.equals(Constants.CNSGNEE_COL) ||
                    val.equals(Constants.WEIGHT_COL) || val.equals(Constants.FREIGHT_NEW_FORMAT_SHP) || val.equals(Constants.FUEL_NEW_FORMAT)) {
                approvedColumnList.add(new Integer(cell.getColumnIndex()));
                System.out.println("added approved column: " + cell.getColumnIndex() + " Cell: " + val);
            }

        }
        return approvedColumnList;
    }


    //TODO better using "while cellIterator.hasNext" ?
    // TODO faster using threads?
    //copies rows from invoice WB sheet to des customer WB sheet
    public static void copyRowsToCustomersWb(Workbook wb, CellAddress customerIdCellAddress, Map<String, Workbook> mapCustomerToWb, Map<String, String> customerNameToFileName) {
        //iterate over all sheets in workbook
        for (int sheetCount = 0; sheetCount < wb.getNumberOfSheets(); sheetCount++) {
            //add rows from main wb to to customers WBs
            Workbook customerWb;
            final Sheet originalSheet = wb.getSheetAt(sheetCount);
            Sheet customerSheet;
            Row originalRow, customerRow;
            String customerName, customerFileName, product;
            final DataFormatter dataFormatter = new DataFormatter();

            Set<Integer> approvedColumnList = approvedColIdList(originalSheet);

            CellAddress cellProduct = findCellByName(originalSheet, Constants.PRODUCTS_COL);
            if (cellProduct == null) {
                throw new IllegalArgumentException(Constants.DES_COL_NOT_FOUND_ERROR);
            }
            int productsCol = cellProduct.getColumn();

            //iterate over all cells in workbook
            for (int i = 1; i <= originalSheet.getLastRowNum(); i++) {
                originalRow = originalSheet.getRow(i);
                customerName = dataFormatter.formatCellValue(originalRow.getCell(customerIdCellAddress.getColumn())); //TODO which is better ? dataFormatter.formatCellValue OR getStringCellValue
                customerFileName = customerNameToFileName.get(customerName);
                customerWb = mapCustomerToWb.get(customerFileName);
                product = originalRow.getCell(productsCol).getStringCellValue();

                //making sure customer exist in map and has a wb file and needs to copy his rows (product=DOX or WPX)
                if (customerFileName != null && customerWb != null && (product.equals(Constants.WPX) || product.equals(Constants.DOX))) {
                    customerSheet = customerWb.getSheet(Constants.SHEET_NAME);
                    //create Row in new sheet
                    int customerRowIndex = customerSheet.getLastRowNum() + 1;
                    customerRow = customerSheet.createRow(customerRowIndex);
                    //copy Rows(cell by cell) to customer sheet
                    for (int j = 0; j < originalRow.getLastCellNum(); j++) {
                        // TODO maybe DO NOT COPY cells of unwanted columns? to avoid deletiong of cells and columns later on
                        //if this list contains the current column id, it's approved and can be copied
                        if (approvedColumnList.contains(new Integer(j))) {
                            Cell cell = customerRow.createCell(j);
                            Cell originalCell = originalRow.getCell(j);

                            //Copy style from old cell and apply to new cell
                            //HSSFCellStyle newCellStyle = workbook.createCellStyle();
                            //newCellStyle.cloneStyleFrom(originalCell.getCellStyle());
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
                        System.out.println("copied Row with index: " + customerRowIndex);
                    }
                } else {
                    System.out.println("Skipped customerName:  " + customerName);
                    continue;
                }
            }
            System.out.println("Finished sheet " + originalSheet.getSheetName());
        }
        System.out.println("Finished iterating workbook ");
        System.out.println("******************************");

    }

    //saves and closes open WB files
    public static void saveAndCloseWbFiles(Map<String, Workbook> mapCustomerFileNameToWb) {
        Workbook wb;
        try {
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                //close and save
                FileOutputStream fileOutCustomer = new FileOutputStream("outputdir/customers/" + customerFileName + ".xls");
                wb = mapCustomerFileNameToWb.get(customerFileName);
                wb.write(fileOutCustomer);
                System.out.println(" Wrote workbook file to disk: " + customerFileName);
                fileOutCustomer.close();
                System.out.println(" Closed file: " + customerFileName + ".xls");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("******************************");
    }

    //saves and closes open WB files
    public static void deleteUnusedSheetsInWb(Map<String, Workbook> mapCustomerFileNameToWb) {
        Workbook wb;
        try {
            //load files
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                FileOutputStream fileOutCustomer = new FileOutputStream("outputdir/customers/" + customerFileName + ".xls");
                wb = mapCustomerFileNameToWb.get(customerFileName);
                wb.removeSheetAt(1); //has a problem here the new sheet has no index
                wb.write(fileOutCustomer);
                System.out.println(" Wrote workbook file to disk: " + customerFileName);
                fileOutCustomer.close();
                System.out.println(" Closed file: " + customerFileName + ".xls");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("******************************");
    }

    //saves and closes open WB files
    public static void insertLogo(Map<String, Workbook> mapCustomerFileNameToWb) {
        Workbook wb;
        Sheet sheet;
        try {
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                wb = mapCustomerFileNameToWb.get(customerFileName);
                sheet = wb.getSheet(Constants.SHEET_NAME);
                //sheet.
                //add picture data to this workbook.
                InputStream is = new FileInputStream("inputdir/logo.png");
                byte[] bytes = IOUtils.toByteArray(is);
                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
                is.close();
                CreationHelper helper = wb.getCreationHelper();
                // Create the drawing patriarch.  This is the top level container for all shapes.
                Drawing drawing = sheet.createDrawingPatriarch();
                //add a picture shape
                ClientAnchor anchor = helper.createClientAnchor();
                //set top-left corner of the picture,
                //subsequent call of Picture#resize() will operate relative to it
                anchor.setCol1(15);
                anchor.setRow1(14);
                Picture pict = drawing.createPicture(anchor, pictureIdx);
                //auto-size picture relative to its top-left corner
                pict.resize();
                System.out.println("Added logo to: " + customerFileName);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("******************************");
    }


    //returns a map of Customer to workbook, creates a new WB file per customer including wb headlines
    public static Map<String, Workbook> createCustomerWbMap(Row firstRow, Map<String, String> customerNameToFileName) {
        final Map<String, Workbook> mapCustomerFileNameWb = new HashMap<String, Workbook>();
        for (String customerFileName : customerNameToFileName.values()) {
            mapCustomerFileNameWb.put(customerFileName, createCustomerWb(firstRow));
            System.out.println("created customer wb with name: " + customerFileName);
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
        Iterator<Cell> cellIterator = firstRow.cellIterator();
        DataFormatter dataFormatter = new DataFormatter();
        Cell cell;
        int colIndx;
        String val;
        while (cellIterator.hasNext()) {
            cell = cellIterator.next();
            colIndx = cell.getColumnIndex();
            val = dataFormatter.formatCellValue(cell);
            if (val.equals(Constants.AWB_COL) || val.equals(Constants.SHIP_DATE_COL) || val.equals(Constants.PRODUCTS_COL) ||
                    val.equals(Constants.ORIGIN_COL) || val.equals(Constants.DES_COL) || val.equals(Constants.PERIOD_COL) ||
                    val.equals(Constants.CUSTOMER_COLUMN_NAME) || val.equals(Constants.REF_NUM_COLUMN_NAME) || val.equals(Constants.CNSGNEE_COL) ||
                    val.equals(Constants.WEIGHT_COL) || val.equals(Constants.FREIGHT_NEW_FORMAT_SHP) || val.equals(Constants.FUEL_NEW_FORMAT)) {
                wbRow.createCell(colIndx).setCellValue(val);
            }
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
    public static CellAddress findCellByName(Sheet hssfSheet, String cellName) {
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
                if (dataFormatter.formatCellValue(cell).equals(cellName)) {
                    CellAddress cellAddress = cell.getAddress();
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
        int productsCol;
        String product;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            //make sure product is DOX or WPX before adding the customer to map
            CellAddress cellProduct = findCellByName(hssfSheet, Constants.PRODUCTS_COL);
            if (cellProduct == null) {
                throw new MissingFormatArgumentException(Constants.PRODUCTS_COLUMN_NOT_FOUND_ERROR);
            }
            productsCol = cellProduct.getColumn();
            product = row.getCell(productsCol).getStringCellValue();
            if (product.equals(Constants.DOX) || product.equals(Constants.WPX)) {
                //TODO get the cust name directly instead of iterating over all
                //TODO example: row.getCell(  findCellByName(hssfSheet, Constants.CUST_NAME_CELL).getColumn())    .getStringCellValue);
                cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //checking if in proper column and in values section, row 0 = header
                    if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
                        String customerNameAsInSheet = cell.getStringCellValue();
                        //fileName = customer name in upper case -minus illegal characters
                        String customerFileName = customerNameAsInSheet.replaceAll(Constants.REGEX_FILTER_UNWANTED_CHARS, " ").toUpperCase();
                        if (!customerNameToFile.containsKey(customerNameAsInSheet)) {
                            customerNameToFile.put(customerNameAsInSheet, customerFileName);
                            System.out.println("Added customer: " + customerNameAsInSheet + " with file name: " + customerFileName + ", Product is: " + product);
                        }
                        break;
                    }
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

    //calculate shipment cost(freight) for each customer based on his personal price list wb
    public static void calcFreightForAllCustomers(Map<String, Workbook> mapCustomerFileNameWb, Map<String, Integer> regionsMap, Double fuelSurcharge) {
        for (String customer : mapCustomerFileNameWb.keySet()) {
            //load customer price list
            final Workbook customerPriceListWb = loadWb(Constants.INPUT_DIR + "/" + Constants.CUSTOMER_PRICE_LISTS + "/" + customer + Constants.XLSX_FILE_ENDING);
            calcCustomerFreight(customer, mapCustomerFileNameWb.get(customer), customerPriceListWb, regionsMap, fuelSurcharge);// TODO check not null
        }
        System.out.println("******************************");
    }

    /**
     * calculates freight cell value according to formula
     * Calculates shipment cost(freight) for a customer based on his personal price list wb
     * Formula: CHARGE = [(customer price per region) X weight]  + [(fuel surcharge% * customer price per region)]
     */
    public static void calcCustomerFreight(String customer, Workbook wb, Workbook customerPriceListWb, Map<String, Integer> regionsMap, Double fuelSurcharge) {
        final Sheet sheet = wb.getSheet(Constants.SHEET_NAME);
        //init cells
        int weightCol, destinationCol, freightCol, productsCol, zone;
        double weight, pricePerWeightAndZone, totalPrice;
        DataFormatter dataFormatter = new DataFormatter();
        Row row;
        Cell cell;
        String country;

        System.out.println("Started calcCustomerFreight for : " + customer);

        //get cell values
        CellAddress cellFrieght = findCellByName(sheet, Constants.FREIGHT);
        if (cellFrieght == null) {
            cellFrieght = findCellByName(sheet, Constants.FREIGHT_NEW_FORMAT_SHP);
        }
        if (cellFrieght == null) {
            throw new IllegalArgumentException(Constants.FREIGHT_NOT_FOUND_ERROR);
        }
        freightCol = cellFrieght.getColumn();

        CellAddress cellWeight = findCellByName(sheet, Constants.WEIGHT_COL);
        if (cellWeight == null) {
            throw new IllegalArgumentException(Constants.WEIGHT_NOT_FOUND_ERROR);
        }
        weightCol = cellWeight.getColumn();

        CellAddress cellDes = findCellByName(sheet, Constants.DES_COL);
        if (cellDes == null) {
            throw new IllegalArgumentException(Constants.DES_COL_NOT_FOUND_ERROR);
        }
        destinationCol = cellDes.getColumn();

        CellAddress cellProduct = findCellByName(sheet, Constants.PRODUCTS_COL);
        if (cellProduct == null) {
            throw new IllegalArgumentException(Constants.DES_COL_NOT_FOUND_ERROR);
        }
        productsCol = cellProduct.getColumn();

        //iterate over all rows in customer workbook ( 1 row = 1 shipment)
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            //get current row
            row = sheet.getRow(i);
            cell = row.getCell(productsCol);
            //this system is only handling WPX and DOX product types of shipments
            if (cell.getStringCellValue() == Constants.WPX || cell.getStringCellValue() == Constants.DOX) {
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
                    System.out.println("Country not found: " + country);
                    throw new IllegalArgumentException(Constants.COUNTRY_CODE_ERROR);
                }

                //calc price according to weight and zone
                pricePerWeightAndZone = getCustomerPrice(customerPriceListWb.getSheet(Constants.SHEET_NAME), weight, zone);
                totalPrice = ((1 + fuelSurcharge) * pricePerWeightAndZone);                                   //fuelScp range:[0 - 1]

                //write updated price value to cell
                cell = row.getCell(freightCol);
                cell.setCellValue(totalPrice);
                System.out.println("Updated cell : " + cell.getAddress() + " with a new FRIEGHT value of: " + totalPrice);
            }
            //any other type has to be done Manually
            else {
                System.out.println(" Customer WB: " + customer + " Row: " + row.getRowNum() + "Product is: " + cell.getStringCellValue() + " To ne done Manually ");
                continue;
            }
            System.out.println("******************************");
        }
    }

    /**
     * gets the price per zone for a given customer
     * //find closest  weight and get price example: weight is 11.5 , base =11K
     */
    public static double getCustomerPrice(Sheet priceListSheet, double weight, int zone) {
        //rows and columns are Zero based
        Row row, nextRow;
        Cell cell;
        double baseWeight, nextStepWeight, additionalPrice, nextStepPrice, remWeight, diffPrice, diffWeight, price = 0;
        final int startRow = 2, endRow = 28, baseWeightCol = 0;

        //go over rows in price list from end to start
        for (int i = endRow; i >= startRow; i++) {
            row = priceListSheet.getRow(i);
            cell = row.getCell(baseWeightCol);
            baseWeight = cell.getNumericCellValue();

            if (weight == baseWeight) {
                price = row.getCell(zone).getNumericCellValue();
                System.out.println("Found exact weight: " + weight + " Price is: " + price);
                break;
            }
            //avoiding overflow of table
            else if ((weight > baseWeight) && (i + 1 <= endRow)) {
                baseWeight = cell.getNumericCellValue();
                nextRow = priceListSheet.getRow(i + 1);
                //find closest weight value
                if ((weight >= cell.getNumericCellValue()) && (weight < priceListSheet.getRow(i + 1).getCell(baseWeightCol).getNumericCellValue())) {
                    price = row.getCell(zone).getNumericCellValue(); //price for base wight
                    nextStepPrice = nextRow.getCell(zone).getNumericCellValue(); //price for next step
                    nextStepWeight = nextRow.getCell(baseWeightCol).getNumericCellValue();
                    diffWeight = nextStepWeight - baseWeight;
                    diffPrice = nextStepPrice - price;
                    remWeight = weight - baseWeight;
                    additionalPrice = diffPrice / diffWeight * remWeight;
                    price = price + additionalPrice;
                    System.out.println("Found  price for weight: " + weight + ", Price is: " + price);
                    break;
                }
            } else {
                continue;
            }
        }
        if (price == 0) {//no price in price list
            System.out.println("weight: " + weight + " not found in price list table ");
            throw new IllegalArgumentException(Constants.WEIGHT_NOT_FOUND_ERROR);
        }
        return price;
    }

    //returns fuel surcharge
    public static double getFuelSurchargePercent() {
        //TODO
        //get getFuelSurchargePercent from ext field
        //if (fuelScp < 0 || fuelScp > 1) {
        //throw new IllegalArgumentException(Constants.FUEL_SURCHARGE_NOT_IN_RANGE);
        //}
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
                    //System.out.println("Found Cell, Column: " + cell.getColumnIndex() + " Row: " + cell.getRowIndex());
                    CellAddress cellAddress = cell.getAddress();
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

    //populates map of Customer names to customer File Names from workbook (without illegal characters
    public static Map<String, String> populateCustomerAndFileNames(Workbook invoiceWb) {
        Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHEET_NUM);
        int customerColName;
        CellAddress cellCustomerName = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
        if (cellCustomerName == null) {
            throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
        }
        customerColName = cellCustomerName.getColumn();
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

//    /**
//     * Given a sheet, this method deletes a column from a sheet and moves
//     * all the columns to the right of it to the left one cell.
//     * <p>
//     * Note, this method will not update any formula references.
//     *
//     * @param sheet //@param column
//     */
//    public static void deleteColumn(Sheet sheet, int columnToDelete) {
//        int maxColumn = 0;
//        for (int r = 0; r < sheet.getLastRowNum() + 1; r++) {
//            Row row = sheet.getRow(r);
//
//            // if no row exists here; then nothing to do; next!
//            if (row == null)
//                continue;
//
//            // if the row doesn't have this many columns then we are good; next!
//            int lastColumn = row.getLastCellNum();
//            if (lastColumn > maxColumn)
//                maxColumn = lastColumn;
//
//            if (lastColumn < columnToDelete)
//                continue;
//
//            for (int x = columnToDelete + 1; x < lastColumn + 1; x++) {
//                Cell oldCell = row.getCell(x - 1);
//                if (oldCell != null)
//                    row.removeCell(oldCell);
//
//                Cell nextCell = row.getCell(x);
//                if (nextCell != null) {
//                    Cell newCell = row.createCell(x - 1, nextCell.getCellType());
//                    cloneCell(newCell, nextCell);
//                }
//            }
//        }
//
//
//        // Adjust the column widths
//        for (int c = 0; c < maxColumn; c++) {
//            sheet.setColumnWidth(c, sheet.getColumnWidth(c + 1));
//        }
//    }
//
//
//    /*
//     * Takes an existing Cell and merges all the styles and forumla
//     * into the new one
//     */
//    private static void cloneCell(Cell cNew, Cell cOld) {
//        cNew.setCellComment(cOld.getCellComment());
//        cNew.setCellStyle(cOld.getCellStyle());
//
//        switch (cNew.getCellType()) {
//            case Cell.CELL_TYPE_BOOLEAN: {
//                cNew.setCellValue(cOld.getBooleanCellValue());
//                break;
//            }
//            case Cell.CELL_TYPE_NUMERIC: {
//                cNew.setCellValue(cOld.getNumericCellValue());
//                break;
//            }
//            case Cell.CELL_TYPE_STRING: {
//                cNew.setCellValue(cOld.getStringCellValue());
//                break;
//            }
//            case Cell.CELL_TYPE_ERROR: {
//                cNew.setCellValue(cOld.getErrorCellValue());
//                break;
//            }
//            case Cell.CELL_TYPE_FORMULA: {
//                cNew.setCellFormula(cOld.getCellFormula());
//                break;
//            }
//        }
//
//    }


    //imported method - does not work 100% well
    public static void deleteColumn(Sheet sheet, int columnToDelete) {
        for (int rId = 0; rId < sheet.getLastRowNum(); rId++) {
            Row row = sheet.getRow(rId);
            for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
                Cell cOld = row.getCell(cID);
                if (cOld != null) {
                    row.removeCell(cOld);
                }
//                Cell cNext = row.getCell(cID + 1);
//                if (cNext != null) {
//                    Cell cNew = row.createCell(cID, cNext.getCellTypeEnum());
//                    cloneCell(cNew, cNext);
//                    //Set the column width only on the first row.
//                    //Other wise the second row will overwrite the original column width set previously.
//                    if(rId == 0) {
//                        sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));
//
//                    }
//                }
            }
        }
    }

    //imported method - does not work 100% well
    public static void cloneCell(Cell cNew, Cell cOld) {
        cNew.setCellComment(cOld.getCellComment());
        cNew.setCellStyle(cOld.getCellStyle());

        if (CellType.BOOLEAN == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getBooleanCellValue());
        } else if (CellType.NUMERIC == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getNumericCellValue());
        } else if (CellType.STRING == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getStringCellValue());
        } else if (CellType.ERROR == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getErrorCellValue());
        } else if (CellType.FORMULA == cNew.getCellTypeEnum()) {
            cNew.setCellValue(cOld.getCellFormula());
        }
    }


//    public static void copyToPdf() {
//
//        FileInputStream input_document = new FileInputStream(new File("C:\\excel_to_pdf.xls"));
//            // Read workbook into HSSFWorkbook
//            HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document);
//            // Read worksheet into HSSFSheet
//            HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
//            // To iterate over the rows
//            Iterator<Row> rowIterator = my_worksheet.iterator();
//            //We will create output PDF document objects at this point
//            Document iText_xls_2_pdf = new Document();
//            PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("Excel2PDF_Output.pdf"));
//            iText_xls_2_pdf.open();
//            //we have two columns in the Excel sheet, so we create a PDF table with two columns
//            //Note: There are ways to make this dynamic in nature, if you want to.
//            PdfPTable my_table = new PdfPTable(2);
//            //We will use the object below to dynamically add new data to the table
//            PdfPCell table_cell;
//            //Loop through rows.
//            while(rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                Iterator<Cell> cellIterator = row.cellIterator();
//                while(cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next(); //Fetch CELL
//                    switch(cell.getCellType()) { //Identify CELL type
//                        //you need to add more code here based on
//                        //your requirement / transformations
//                        case Cell.CELL_TYPE_STRING:
//                            //Push the data from Excel to PDF Cell
//                            table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));
//                            //feel free to move the code below to suit to your needs
//                            my_table.addCell(table_cell);
//                            break;
//                    }
//                    //next line
//                }
//
//            }
//            //Finally add the table to PDF document
//            iText_xls_2_pdf.add(my_table);
//            iText_xls_2_pdf.close();
//            //we created our pdf file..
//            input_document.close(); //close xls
//        }


}
