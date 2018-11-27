package Controller;

public class OldMethods {




// - DEPRECATED METHODS

    //prints workbook info (sheet names and numbers)
//    public static void printWorkbookInfo(Workbook workbook) {
//        //Headlines
//        System.out.println(" Workbook Data: ");
//        System.out.println(" ------------------------------------ ");
//
//        //workbook
//        //String workbookName = workbook.getName();
//        //System.out.println(" Workbook name is: " + workbookName);
//
//        //sheets in workbook
//        int numOfSheets = workbook.getNumberOfSheets();
//        System.out.println("in workbook: " + "there are: " + numOfSheets + " sheets ");
//
//        //loop through sheets
//        for (int i = 0; i < numOfSheets; i++) {
//            System.out.println(" Sheet name: " + workbook.getSheetName(i) + " , Sheet number: " + i);
//            Sheet sheet = workbook.getSheetAt(i);
//        }
//    }

//    /**
//     * copy customer rows from main sheet to customer sheets
//     * iterate over rows ,for each row:
//     * get cust name
//     * if cust xls exist 	-> copy row to cust xls
//     * else 				-> create new xls, new sheet, copy row to cust xls
//     * copy row needs to check that the destination row does not exist before creating it
//     * save & close all files
//     *
//     * @param mainSheet
//     * @param customerNameColumn
//     * @param customersSet
//     */
//    public static void copyCustomerRowsToCustomerSheet(Sheet mainSheet, int customerNameColumn, Set<String> customersSet) {
//        //iterate over sheet rows
//        Iterator<Row> rowIterator = mainSheet.iterator();
//        //loop through rows in sheet
//        while (rowIterator.hasNext()) {
//
//            //Get the row object
//            Row row = rowIterator.next();
//            int rowNum = row.getRowNum();
//            if (rowNum == 0) {
//                continue;
//            }
//
//            //Every row has columns, get the column iterator and iterate over them
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            //check customer
//            String customerName = row.getCell(customerNameColumn).getStringCellValue();// TODO change {getStringCellValue} to {dataFormatter.formatCellValue()} ?
//
//            String pathToFile = "outputdir/customers/workbook " + customerName;
//            try {
//                Workbook customerWorkbook = loadWb(pathToFile);
//                Sheet customerSheet = customerWorkbook.getSheet(customerName);
//                Row customerRow = customerSheet.createRow(rowNum);
//                //copy rows:
//                //copyRows(row, customerRow); // TODO come back here later
//                String name = "";
//                String shortCode = "";
//                //loop through cells
//                while (cellIterator.hasNext()) {
//                    //Get the Cell object
//                    Cell cell = cellIterator.next();
//                    //check the cell type and process accordingly
//                    switch (cell.getCellType()) {
//                        case Cell.CELL_TYPE_STRING:
//                            if (shortCode.equalsIgnoreCase("")) {
//                                shortCode = cell.getStringCellValue().trim();// TODO change {getStringCellValue} to {dataFormatter.formatCellValue()} ?
//                            } else if (name.equalsIgnoreCase("")) {
//                                //2nd column
//                                name = cell.getStringCellValue().trim();
//                            } else {
//                                //random data, leave it
//                                System.out.println("Random data::" + cell.getStringCellValue());
//                            }
//                            break;
//                        case Cell.CELL_TYPE_NUMERIC:
//                            System.out.println("Random data::" + cell.getNumericCellValue());
//                    }
//                } //end of cell iterator
//
//            } catch (Exception e) {
//                e.printStackTrace(printStream);
//            }
//
//
//        } //end of rows iterator
//    }

    //copies ROW
//    private static void copyRow(Workbook workbook, Sheet worksheet, Sheet resultSheet, int sourceRowNum, int destinationRowNum) {
//        Row newRow = resultSheet.getRow(destinationRowNum);
//
//        Row sourceRow = worksheet.getRow(sourceRowNum);
//
//        // If the row exist in destination, push down all rows by 1 else create a new row
//        if (newRow != null) {
//            resultSheet.shiftRows(destinationRowNum, resultSheet.getLastRowNum(), 1);
//        } else {
//            newRow = resultSheet.createRow(destinationRowNum);
//        }
//
//        // Loop through source columns to add to new row
//        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
//            // Grab a copy of the old/new cell
//            Cell oldCell = sourceRow.getCell(i);
//            Cell newCell = newRow.createCell(i);
//
//            // If the old cell is null jump to next cell
//            if (oldCell == null) {
//                newCell = null;
//                continue;
//            }
//
//            // Copy style from old cell and apply to new cell
//            CellStyle newCellStyle = workbook.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            newCell.setCellStyle(newCellStyle);
//
//            // If there is a cell comment, copy
//            if (oldCell.getCellComment() != null) {
//                newCell.setCellComment(oldCell.getCellComment());
//            }
//
//            // If there is a cell hyperlink, copy
//            if (oldCell.getHyperlink() != null) {
//                newCell.setHyperlink(oldCell.getHyperlink());
//            }
//
//            // Set the cell data type
//            newCell.setCellType(oldCell.getCellType());
//
//            // Set the cell data value
//            switch (oldCell.getCellType()) {
//                case Cell.CELL_TYPE_BLANK:
//                    newCell.setCellValue(oldCell.getStringCellValue());
//                    break;
//                case Cell.CELL_TYPE_BOOLEAN:
//                    newCell.setCellValue(oldCell.getBooleanCellValue());
//                    break;
//                case Cell.CELL_TYPE_ERROR:
//                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
//                    break;
//                case Cell.CELL_TYPE_FORMULA:
//                    newCell.setCellFormula(oldCell.getCellFormula());
//                    break;
//                case Cell.CELL_TYPE_NUMERIC:
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                    break;
//                case Cell.CELL_TYPE_STRING:
//                    newCell.setCellValue(oldCell.getRichStringCellValue());
//                    break;
//            }
//        }
//    }

    //returns list of columns strings
//    public static List<String> getStringListOfColumns(Sheet hssfSheet) {
//        //list of columns
//        List<String> columnList = new ArrayList<String>();
//        //row iterator
//        Iterator<Row> rowIterator = hssfSheet.iterator();
//        Iterator<Cell> cellIterator;
//        Row row;
//        Cell cell;
//        while (rowIterator.hasNext()) {
//            //get line
//            row = rowIterator.next();
//            while (row.cellIterator().hasNext()) {
//                //get cell
//                cellIterator = row.cellIterator();
//                cell = cellIterator.next();
//                //add to list
//                columnList.add(cell.getStringCellValue());
//                //get cell address
//                cell.getAddress();
//                //get column index
//                cell.getColumnIndex();
//            }
//
//        }
//        //
//        return columnList;
//    }

    //prints data from workbook
//    public static void printDataFromWorkbook(Workbook workbook) throws Exception {
//
//        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//            Sheet hssfSheet = workbook.getSheetAt(i);
//            System.out.println(" Sheet name: " + hssfSheet.getSheetName() + " , Sheet number: " + i);
//            printDataFromSheet(hssfSheet);
//        }
//    }

    //finds customer name column
//    public static CellRangeAddress findCellRangeAddress(Sheet hssfSheet, CellAddress customerNameCellAddress) throws ClassNotFoundException {
//        CellRangeAddress cellRangeAddress;
//        int firstRow, lastRow, firstCol, lastCol;
//        boolean firstRowFlag, lastRowFlag, firstColFlag, lastColFlag;
//
//        //row iterator
//        Iterator<Row> rowIterator = hssfSheet.iterator();
//        Iterator<Cell> cellIterator;
//        Row row;
//        Cell cell;
//        while (rowIterator.hasNext()) {
//            //get line
//            row = rowIterator.next();
//
//            cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                //get cell
//                cell = cellIterator.next();
//                //check if cell string is customer name column
//                if (cell.getStringCellValue().equals(Constants.CUSTOMER_COLUMN_NAME)) {
//                    //System.out.println("Found Cell, Column: " + cell.getColumnIndex() + " Row: " + cell.getRowIndex());
//                    CellAddress cellAddress = cell.getAddress();
//                    return null;
//                }
//            }
//        }
//        throw new ClassNotFoundException(Constants.COLUMN_NOT_FOUND_ERROR);
//    }

    //creates workbook
//    public static Workbook createWorkbook() throws Exception {
//        //Create Blank Workbook
//        XSSFWorkbook workbook = new XSSFWorkbook();
//        //Create file system using specific name
//        FileOutputStream out = new FileOutputStream(new File("outputdir/createWorkbook.xlsx"));
//        //write operation Workbook using file out object
//        workbook.write(out);
//        out.close();
//        System.out.println("createWorkbook.xlsx written successfully");
//        return workbook;
//    }

    //gets customer names from workbook
//    public static HashSet<String> getCustomerNamesFromWorkbook(Workbook workbook, CellAddress cellAddress) {
//        List<String> customerList = new ArrayList();
//        int numOfSheets = workbook.getNumberOfSheets();
//        Sheet hssfSheet;
//        int columnIndex = cellAddress.getColumn();
//        for (int i = 0; i < numOfSheets; i++) {
//            hssfSheet = (Sheet) workbook.getSheetAt(i);
//            //customerList.addAll(getCustomerNamesAndFileNamesFromSheet(hssfSheet, columnIndex));
//        }
//        return new HashSet<String>(customerList);
//    }

    //clean illegal chars
//    public static void removeIllegalCharactersFromColumnInSheet(Sheet sheet, int columnIndex) {
//        Iterator<Row> rowIterator = sheet.iterator();
//        Iterator<Cell> cellIterator;
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//            cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                //checking if in proper column and in values section
//                if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
//                    //remove un-needed characters from name
//                    cell.setCellValue(cell.getStringCellValue().replaceAll("[\\-\\+\\.\\^:,/]", ""));
//                    break;
//                }
//            }
//        }
//    }

    //gets customer IDs from workbook
//    public static HashSet<Customer> getCustomerIdsFromWorkbook(Workbook workbook, CellAddress customerIdCellAddress, Map<String, String> customerIdToNameMap) {
//        final int numOfSheets = workbook.getNumberOfSheets();
//        final int columnIndex = customerIdCellAddress.getColumn();
//        final HashSet<String> customerIdsSet = new HashSet<String>();
//        final HashSet<Customer> customerSet = new HashSet<Customer>();
//        Sheet sheet;
//
//        //iterate over sheets on wb, for each sheet clean, validate and add customer IDs to List
//        for (int i = 0; i < numOfSheets; i++) {
//            sheet = workbook.getSheetAt(i);
//            customerIdsSet.addAll(getCustomerIdsFromSheet(sheet, columnIndex));
//        }
//
//        //create set of customers from Set of customer IDs(enrich)
//        for (String id : customerIdsSet) {
//            customerSet.add(new Customer(id, customerIdToNameMap.get(id)));
//        }
//        return customerSet;
//    }

    //read customer ids file, returns ID to Name map
//    public static Map<String, String> loadCustomerIdToNameMap(Sheet sheet, int idColumn, int nameColumn) {
//        Row row;
//        String customerId;
//        String customerName;
//        Map<String, String> idToNameMap = new HashMap<String, String>();
//        DataFormatter dataFormatter = new DataFormatter();
//
//        //iterate over all cells in workbook
//        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
//            row = sheet.getRow(i);
//            //validate format, change names
//            customerId = dataFormatter.formatCellValue(row.getCell(idColumn)).replaceAll(Constants.REGEX_ONLY_NUMBERS, "");
//            //TODO improve reg-ex to include only a-z chars and numbers
//            customerName = dataFormatter.formatCellValue(row.getCell(nameColumn)).replaceAll(Constants.REGEX_ILLEGAL_CHARS, "");
//            if (!idToNameMap.containsKey(customerId)) {
//                idToNameMap.put(customerId, customerName);
//            }
//        }
//        return idToNameMap;
//    }

    //creates(copies an example) a price list file per customer in the input workbook,
    //example call:         UtilityMethods.createPriceListFiles(new ArrayList<String>(customerNameToFileName.values()));
//    public static void createPriceListFiles(List<String> names) {
//        //create files price list per customer
//        File folder = new File("inputdir/customer price lists/");
//        List<String> list = Arrays.asList(folder.list());
//        for (String s : names) {
//            if (!list.contains(s + " " + Constants.PL_FILE_ENDING)) {
//                copyFile("inputdir/customer price lists/EXAMPLE CUSTOMER PRICE LIST.xlsx", "inputdir/customer price lists/" + s + Constants.XLSX_FILE_ENDING);
//                System.out.println("created file for: " + s);
//            } else {
//                System.out.println("file: " + s + " Already exists");
//            }
//        }
//    }

    //populates map of Customer names to customer File Names from workbook (without illegal characters
//    public static Map<String, String> populateCustomerAndFileNames(Workbook invoiceWb) {
//        Sheet invoiceSheet = invoiceWb.getSheetAt(Constants.FIRST_SHEET_NUM);
//        int customerColName;
//        CellAddress cellCustomerName = UtilityMethods.findCellByName(invoiceSheet, Constants.CUSTOMER_COLUMN_NAME);
//        if (cellCustomerName == null) {
//            throw new IllegalArgumentException(Constants.CUSTOMER_COLUMN_NOT_FOUND_ERROR);
//        }
//        customerColName = cellCustomerName.getColumn();
//        final Map<String, String> customerNameToFileName = new HashMap<String, String>();
//        //get customer names from invoice wb and map to file names
//        for (int i = 0; i < invoiceWb.getNumberOfSheets(); i++) {
//            invoiceSheet = invoiceWb.getSheetAt(i);
//            customerNameToFileName.putAll(UtilityMethods.getCustomerNamesAndFileNamesFromSheet(invoiceSheet, customerColName, customerNameToFileName)); //TODO using put all could be a problem when having same cust on multiple sheets could ovveride it and create multiple files
//        }
//        return customerNameToFileName;
//    }

    //log and report method to check if each customer in the map has a respective price list file
//    public static void printPriceListFilesInfo(Map<String, String> customerNameToFileName) {
//        //get names of customer price list folder
//        final File folder = new File(Constants.INPUT_DIR + "/" + Constants.CUSTOMER_PRICE_LISTS + "/");
//        final File[] files = folder.listFiles();
//        final Set<String> priceListFileNames = new HashSet<String>();
//        System.out.println("******************************");
//        System.out.println("priceListFiles... ");
//        for (File file : files) {
//            priceListFileNames.add(file.getName());
//            System.out.println("price List File: " + file.getName());
//        }
//        System.out.println("******************************");
//
//        //check if there are price lists per customer and log
//        System.out.println("******************************");
//        System.out.println("customer names not in price list ");
//        for (String fileName : customerNameToFileName.values()) {
//            if (!priceListFileNames.contains(fileName + Constants.XLSX_FILE_ENDING)) {
//                System.out.println(fileName + " is not in priceListFileNames");
//            }
//        }
//        System.out.println("******************************");
//    }

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
//    public static void deleteColumn(Sheet sheet, int columnToDelete) {
//        for (int rId = 0; rId < sheet.getLastRowNum(); rId++) {
//            Row row = sheet.getRow(rId);
//            for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
//                Cell cOld = row.getCell(cID);
//                if (cOld != null) {
//                    row.removeCell(cOld);
//                }
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
//            }
//        }
//    }

    //imported method - does not work 100% well
//    public static void cloneCell(Cell cNew, Cell cOld) {
//        cNew.setCellComment(cOld.getCellComment());
//        cNew.setCellStyle(cOld.getCellStyle());
//
//        if (CellType.BOOLEAN == cNew.getCellTypeEnum()) {
//            cNew.setCellValue(cOld.getBooleanCellValue());
//        } else if (CellType.NUMERIC == cNew.getCellTypeEnum()) {
//            cNew.setCellValue(cOld.getNumericCellValue());
//        } else if (CellType.STRING == cNew.getCellTypeEnum()) {
//            cNew.setCellValue(cOld.getStringCellValue());
//        } else if (CellType.ERROR == cNew.getCellTypeEnum()) {
//            cNew.setCellValue(cOld.getErrorCellValue());
//        } else if (CellType.FORMULA == cNew.getCellTypeEnum()) {
//            cNew.setCellValue(cOld.getCellFormula());
//        }
//    }


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
