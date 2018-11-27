package Controller;

import Model.Constants;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class UtilityMethods {

    //used to direct system out into a printed stream
    private static PrintStream printStream;

    //c'tor
    public UtilityMethods(PrintStream printStream) {
        this.printStream = printStream;
    }

    //sets page print setup and closes open WB files
    public static void pagePrintSetup(Map<String, Workbook> mapCustomerFileNameToWb) {
        Workbook wb;
        Sheet sheet;
        try {
            //load files
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                wb = mapCustomerFileNameToWb.get(customerFileName);
                sheet = wb.getSheet(Constants.SHEET_NAME);
                sheet.setFitToPage(true);
                sheet.getPrintSetup().setLandscape(true);
                sheet.getPrintSetup().setFitHeight((short) 1);
                sheet.getPrintSetup().setFitWidth((short) 1);
            }
        } catch (Exception e) {
            e.printStackTrace(printStream);
        }
        //System.out.println("******************************");
    }


    //returns a list of approved column ids, approved = viewable by customer
    public static Set<Integer> approvedColIdList(Sheet sheet) {
        final Set<Integer> approvedColumnList = new HashSet<Integer>();
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
                //System.out.println("added approved column: " + cell.getColumnIndex() + " Cell: " + val);
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
                customerName = dataFormatter.formatCellValue(originalRow.getCell(customerIdCellAddress.getColumn()));
                customerFileName = customerNameToFileName.get(customerName);
                customerWb = mapCustomerToWb.get(customerFileName);
                product = dataFormatter.formatCellValue(originalRow.getCell(productsCol));

                //making sure customer exist in map and has a wb file and needs to copy his rows (product=DOX or WPX)
                if (customerFileName != null && customerWb != null && (product.equals(Constants.WPX) || product.equals(Constants.DOX))) {
                    customerSheet = customerWb.getSheet(Constants.SHEET_NAME);
                    //create Row in new sheet
                    int customerRowIndex = customerSheet.getLastRowNum() + 1;
                    customerRow = customerSheet.createRow(customerRowIndex);
                    //copy Rows(cell by cell) to customer sheet
                    for (int j = 0; j < originalRow.getLastCellNum(); j++) {
                        //if this list contains the current column id, it's approved and can be copied
                        if (approvedColumnList.contains(new Integer(j))) {
                            Cell cell = customerRow.createCell(j);
                            Cell originalCell = originalRow.getCell(j);

                            //TODO
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
                    }
                } else {
                    continue;
                }
            }
        }
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
                fileOutCustomer.close();
            }
        } catch (Exception e) {
            e.printStackTrace(printStream);
        }
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
            e.printStackTrace(printStream);
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
            e.printStackTrace(printStream);
        }
        System.out.println("******************************");
    }


    //returns a map of Customer to workbook, creates a new WB file per customer including wb headlines
    public static Map<String, Workbook> createCustomerWbMap(Row firstRow, Map<String, String> customerNameToFileName) {
        final Map<String, Workbook> mapCustomerFileNameWb = new HashMap<String, Workbook>();
        for (String customerFileName : customerNameToFileName.values()) {
            mapCustomerFileNameWb.put(customerFileName, createCustomerWb(firstRow));
            //System.out.println("created customer wb with name: " + customerFileName);
        }
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
                            shortCode = cell.getStringCellValue().trim(); // TODO change {getStringCellValue} to {dataFormatter.formatCellValue()} ?
                        } else if (name.equalsIgnoreCase("")) {
                            //2nd column
                            name = cell.getStringCellValue().trim();// TODO change {getStringCellValue} to {dataFormatter.formatCellValue()} ?
                        } else {
                            //random data, leave it
                            System.out.println("Random data::" + cell.getStringCellValue());// TODO change {getStringCellValue} to {dataFormatter.formatCellValue()} ?
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
                        //System.out.println("old Cell Value: " + cellVal + " changed to new cell value: " + cellVal.replaceAll(Constants.REGEX_ONLY_NUMBERS, Constants.BLANK));
                        cell.setCellValue(cellVal.replaceAll(Constants.REGEX_ONLY_NUMBERS, Constants.BLANK)); //TODO correct place to change cell values ?
                    }
                    customerIdsSet.add(dataFormatter.formatCellValue(cell));
                    break;
                }
            }
        }
        return customerIdsSet;
    }

    //gets customer names from a sheet, "cleans" names by removing from unwanted characters
    public static Map<String, String> getCustomerNamesAndFileNamesFromSheet(Sheet hssfSheet, int columnIndex, Map<String, String> customerNameToFile) {
        if (customerNameToFile == null) {
            throw new NullPointerException(Constants.MAP_IS_NULL);
        }
        final Iterator<Row> rowIterator = hssfSheet.iterator();
        Iterator<Cell> cellIterator;
        int productsCol;
        String product;
        final DataFormatter dataFormatter = new DataFormatter();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            //make sure product is DOX or WPX before adding the customer to map
            CellAddress cellProduct = findCellByName(hssfSheet, Constants.PRODUCTS_COL);
            if (cellProduct == null) {
                throw new MissingFormatArgumentException(Constants.PRODUCTS_COLUMN_NOT_FOUND_ERROR);
            }
            productsCol = cellProduct.getColumn();
            if (row.getCell(productsCol) == null) {
                continue;
            }
            product = dataFormatter.formatCellValue(row.getCell(productsCol));
            if (product.equals(Constants.DOX) || product.equals(Constants.WPX)) {
                //TODO get the cust name directly instead of iterating over all
                //TODO example: row.getCell(  findCellByName(hssfSheet, Constants.CUST_NAME_CELL).getColumn())    .getStringCellValue);
                cellIterator = row.cellIterator();
                String customerNameAsInSheet;
                String customerFileName;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //checking if in proper column and in values section, row 0 = header
                    if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
                        customerNameAsInSheet = dataFormatter.formatCellValue(cell);
                        //fileName = customer name in upper case -minus illegal characters
                        customerFileName = customerNameAsInSheet.replaceAll(Constants.REGEX_FILTER_UNWANTED_CHARS, "").replaceAll(Constants.REGEX_REPLACE_2_SPACES_WITH_1, " ").trim().toUpperCase();
                        if (!customerNameToFile.containsKey(customerNameAsInSheet)) {
                            customerNameToFile.put(customerNameAsInSheet, customerFileName);
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
        } catch (IOException e) {
            e.printStackTrace(printStream);
        }
    }

    //deletes file
    public static void deleteFile(String file) {
        try {
            FileUtils.deleteQuietly(new File(file));
        } catch (Exception e) {
            e.printStackTrace(printStream);
        }
    }

    //loads workbook from path to file
    public static Workbook loadWb(String path) {
        //read file contents
        if (path == null || path.equals("")) {
            return null;
        }
        final File file = new File(path);
        if (file == null || !file.exists()) {
            return null;
        }
        try {
            final FileInputStream fIP = new FileInputStream(file);
            //Get the Workbook instance for XLS file
            final Workbook workbook = WorkbookFactory.create(fIP);
            return workbook;
        } catch (Exception e) {
            e.printStackTrace(printStream);
        }
        return null;
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
        Workbook customerPriceListWb = null;
        for (String customer : mapCustomerFileNameWb.keySet()) {
            //load customer price list
            try {
                customerPriceListWb = loadWb(Constants.INPUT_DIR + "/" + Constants.CUSTOMER_PRICE_LISTS + "/" + customer + Constants.XLSX_FILE_ENDING);
            } catch (Exception e) {
                e.printStackTrace(printStream);
            }
            if (customerPriceListWb == null) {
                System.out.println(" customer price list file not found for: " + customer + " needs to be done Manually");
                continue;
            }
            calcCustomerFreight(customer, mapCustomerFileNameWb.get(customer), customerPriceListWb, regionsMap, fuelSurcharge);
        }
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
        String country, cellStringValue;

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

        //iterate over all rows in customer workbook ( 1x row = 1x shipment)
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            //get current row
            row = sheet.getRow(i);
            cell = row.getCell(productsCol);
            cellStringValue = dataFormatter.formatCellValue(cell);
            //this system is only handling WPX and DOX product types of shipments
            if (cellStringValue.equals(Constants.WPX) || cellStringValue.equals(Constants.DOX)) {
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
                    throw new IllegalArgumentException(Constants.COUNTRY_CODE_ERROR);
                }

                //calc price according to weight and zone
                pricePerWeightAndZone = getCustomerPrice(customerPriceListWb.getSheet(Constants.SHEET_NAME), weight, zone);
                totalPrice = ((1 + fuelSurcharge) * pricePerWeightAndZone);                                   //fuelScp range:[0 - 1]

                //write updated price value to cell
                cell = row.getCell(freightCol);
                cell.setCellValue(totalPrice);
            }
            //any other type has to be done Manually
            else {
                System.out.println(" Customer WB: " + customer + " Row: " + row.getRowNum() + "Product is of type: " + cellStringValue + "Needs To be done Manually ");
                continue;
            }
        }
    }

    /**
     * calculates and returns the customer price per shipment according to the parameters (weight, zone, price list)
     * gets the price per zone for a given customer
     */
    public static double getCustomerPrice(Sheet priceListSheet, double weight, int zone) {
        //rows and columns are Zero based
        Row row, nextRow;
        Cell cell;
        double rowWeight, nextStepWeight, additionalPrice, nextStepPrice, remWeight, diffPrice, diffWeight, price = 0;
        final int startRow = Constants.FIRST_ROW_NUM_PRICE_LIST, endRow = priceListSheet.getLastRowNum(), baseWeightCol = 0;
        int i = endRow;
        try {
            //go over rows in price list from end to start
            while (i >= startRow) {     //TODO impl better search method here? Bin search?
                //get row data
                row = priceListSheet.getRow(i);
                cell = row.getCell(baseWeightCol);
                rowWeight = cell.getNumericCellValue();

                //weight bigger than max weight in price list
                if (weight > rowWeight) {
                    //calc max weight
                    price = row.getCell(zone).getNumericCellValue();
                    diffWeight = weight - rowWeight;
                    additionalPrice = diffWeight * price;
                    price = price + additionalPrice;
                    break;
                }
                //weight equals to row current weight
                if (weight == rowWeight) {
                    price = row.getCell(zone).getNumericCellValue();
                    break;
                }
                //not to under flow
                if (i - 1 >= startRow) {
                    //get next row data
                    nextRow = priceListSheet.getRow(i - 1);
                    nextStepWeight = nextRow.getCell(baseWeightCol).getNumericCellValue();

                    //weight smaller than current and next row weight, jump to next row
                    if (weight < rowWeight && weight <= nextStepWeight) {
                        i--;
                        continue;
                    }
                    //weight is in between current and next row weight
                    if (weight < rowWeight && weight > nextStepWeight) {
                        //calc base price
                        price = row.getCell(zone).getNumericCellValue();
                        //calc price per kg for "in-between" row
                        nextStepPrice = nextRow.getCell(zone).getNumericCellValue(); //price for next step
                        diffWeight = rowWeight - nextStepWeight;
                        diffPrice = price - nextStepPrice;
                        remWeight = weight - rowWeight;
                        additionalPrice = diffPrice / diffWeight * remWeight;
                        price = price + additionalPrice;
                        break;
                    }
                }
            }
            if (price == 0) {//no price in price list
                //System.out.println("weight: " + weight + " not found in price list table ");
                throw new IllegalArgumentException(Constants.WEIGHT_NOT_FOUND_ERROR);
            }
        } catch (Exception e) {
            e.printStackTrace(printStream);
            throw e;
        }
        return price;
    }
}