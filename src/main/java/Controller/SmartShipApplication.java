package Controller;

import Model.Constants;
import Model.Customer;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by ARIELPE on 4/2/2017.
 */
public class SmartShipApplication {


	public static void main(String[] args) {

		try {
			//file paths
			final String invoiceFileSourcePath = "inputdir/DHL Invoices/March - DHL invoice - with customerIds for test.xls";
			final String invoiceFileSheetName = "TLV0000492709"; // used for copying the headline TODO: add this as part of the GLUFA
			final String invoiceFileDestinationPath = "outputdir/invoiceFile workbook" + System.currentTimeMillis() + ".xls";

			final String customerIdsFileSourcePath = "inputdir/customer names to ids mapping/customersNamesToIdsMapping.xls"; // customer names to ID mapping file
			final String customerIdsFileSheetNameCustomerIds = "Customer Names Type and Serials"; //sheet name
			final String customerIdsFileDestinationPath = "outputdir/customerIds workbook " + System.currentTimeMillis() + ".xls";

			final String pathRegionToCountryFile = "inputdir/Region to country map/Regions to Country Mapping.xls";
			final Workbook countryToRegionCodeWb = WorkbookFactory.create(new File(pathRegionToCountryFile)); //TODO instantiate with WorkbookFactory
			final Sheet sheet = countryToRegionCodeWb.getSheet(Constants.SHEET_NAME);
			int regionIndexCol = findCellByName(sheet, Constants.ZONE_NUM_COL).getColumn();
			int regionNameCol = findCellByName(sheet, Constants.COUNTRY_COL).getColumn();
			final Map<String, Integer> countryToRegionCodeMap = loadRegionToCountryMap(sheet, regionIndexCol, regionNameCol);

			//copy customerIds file not to work on original file
			copyFile(customerIdsFileSourcePath, customerIdsFileDestinationPath);

			//load customer Ids workbook TODO slow?
			final Workbook customerIdsWorkbook = loadWorkbook(customerIdsFileDestinationPath);

			//get first sheet of workbook
			final Sheet customerIdSheet = customerIdsWorkbook.getSheet(customerIdsFileSheetNameCustomerIds);

			//create customer  ID to name map
			final Map<String, String> customerIdToNameMap = loadCustomerIdToNameMap(customerIdSheet, Constants.CUST_ID_COL, Constants.CUST_NAME_COL);

			//copy invoice file not to work on original file
			copyFile(invoiceFileSourcePath, invoiceFileDestinationPath);

			//load copied invoice workbook
			final Workbook invoiceWorkbook = loadWorkbook(invoiceFileDestinationPath);

			//get first sheet of workbook
			final Sheet invoiceSheet = invoiceWorkbook.getSheet(invoiceFileSheetName);

			//find cell address of customer reference(id)
			final CellAddress customerIdCellAddress = findCellByName(invoiceSheet, Constants.REF_NUM_COLUMN_NAME);

			//get customer Ids from workbook
			final HashSet<Customer> customerSet = getCustomerIdsFromWorkbook(invoiceWorkbook, customerIdCellAddress, customerIdToNameMap);

			//create customer workbooks,  file for each customer //TODO develope to be able to run on a WB and not only on a sheet using for loop iteration
			final Map<Customer, Workbook> mapCustomerToWorkbook = getCustomerToWbMap(invoiceSheet, customerSet);

			//copy rows to customer workbooks //TODO slow - make faster
			copyRowsToCustomersWb(invoiceWorkbook, customerIdCellAddress, mapCustomerToWorkbook);

			//calc freight cell value from formula per customer
			calcFreightForAllCustomers(mapCustomerToWorkbook, countryToRegionCodeMap);

			saveAndCloseWbAndFiles(mapCustomerToWorkbook);

			System.out.println(" Finished Main ");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	//copies rows from src WB to des customer WB sheets //TODO slow - make faster
	public static void copyRowsToCustomersWb(Workbook workbook, CellAddress customerIdCellAddress, Map<Customer, Workbook> mapCustomerToWorkbook) throws IOException {
		//iterate over all sheets in workbook
		for (int sheetCount = 0; sheetCount < workbook.getNumberOfSheets(); sheetCount++) {
			//add rows from main wb to to customers WBs
			final Sheet originalSheet = workbook.getSheetAt(sheetCount);
			Row originalRow;
			Workbook customerWorkbook;
			Sheet customerSheet;
			Row customerRow;
			final DataFormatter dataFormatter = new DataFormatter();
			//iterate over all cells in workbook
			for (int i = 1; i <= originalSheet.getLastRowNum(); i++) {
				originalRow = originalSheet.getRow(i);
				String customerId = dataFormatter.formatCellValue(originalRow.getCell(customerIdCellAddress.getColumn()));

				//get current customer wb, sheet
				customerWorkbook = mapCustomerToWorkbook.get(new Customer(customerId));
				customerSheet = customerWorkbook.getSheet(Constants.SHEET_NAME);

				//create originalRow and copy content
				int customerRowIndex = customerSheet.getLastRowNum() + 1;
				customerRow = customerSheet.createRow(customerRowIndex);

				//copy cells from originalRow to customer sheet, in a new row
				for (int j = 0; j < originalRow.getLastCellNum(); j++) {    //TODO better to use {while cellIterator.hasNext()} ?
					Cell cell = customerRow.createCell(j);
					Cell originalCell = originalRow.getCell(j);

//					Copy style from old cell and apply to new cell
//					HSSFCellStyle newCellStyle = workbook.createCellStyle();
//					newCellStyle.cloneStyleFrom(originalCell.getCellStyle()); //TODO fix styling of cells later
//					cell.setCellStyle(newCellStyle);

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
	}

	//saves and closes open WB files
	public static void saveAndCloseWbAndFiles(Map<Customer, Workbook> mapCustomerToWorkbook) throws IOException {
		//save and close workbooks and files
		for (Customer customer : mapCustomerToWorkbook.keySet()) {
			//close and save
			FileOutputStream fileOutCust = new FileOutputStream("outputdir/customers/" + customer.getName() + ".xls");
			mapCustomerToWorkbook.get(customer).write(fileOutCust);
			System.out.println(" Wrote workbook to disk: " + customer.getName() );
			fileOutCust.close();
			System.out.println(" Closed file: " + customer.getName() + ".xls");
		}
	}

	//returns a map of Customer to workbook, creates a new WB file per customer including wb headlines
	public static Map<Customer, Workbook> getCustomerToWbMap(Sheet hssfSheet, HashSet<Customer> customerSet) {
		Map<Customer, Workbook> mapCustomerWorkbook = new HashMap<Customer, Workbook>();
		for (Customer customer : customerSet) {
			Workbook wb = new HSSFWorkbook();
			Sheet wbSheet = wb.createSheet(Constants.SHEET_NAME);
			Row wbRow = wbSheet.createRow(Constants.FIRST_ROW_NUM);
			//fill row 0 with headlines
			final Row firstRow = hssfSheet.getRow(0);
			for (int cellCount = 0; cellCount < firstRow.getLastCellNum(); cellCount++) {
				wbRow.createCell(cellCount).setCellValue(firstRow.getCell(cellCount).getStringCellValue());
			}
			mapCustomerWorkbook.put(customer, wb);
			System.out.println("created customer wb with name: " + customer.getName() + " and id: " + customer.getId());
		}
		return mapCustomerWorkbook;
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
				Workbook customerWorkbook = loadWorkbook(pathToFile);
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

	//finds cell address by name
	public static CellAddress findCellByName(Sheet hssfSheet, String customerColumnName) throws ClassNotFoundException {
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
		throw new ClassNotFoundException(Constants.COLUMN_NOT_FOUND_ERROR);
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
			customerId = dataFormatter.formatCellValue(row.getCell(idColumn));
			//validate format, change names
			customerName = dataFormatter.formatCellValue(row.getCell(nameColumn)).replaceAll(Constants.REGEX_ILLEGAL_CHARS, "");;
			//TODO: fix reg ex here
//			if (customerName.contains(Constants.REGEX_ILLEGAL_CHARS)){
//				System.out.println("updated customer name from: " + customerName + "to " + customerName.replaceAll(Constants.REGEX_ONLY_NUMBERS_A_TO_Z_LETTERS, Constants.BLANK));
//				customerName = customerName.replaceAll(Constants.REGEX_ILLEGAL_CHARS, "");
//			}
			idToNameMap.put(customerId, customerName);
		}
		return idToNameMap;
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

	//gets customer names from workbook
	public static HashSet<String> getCustomerNamesFromWorkbook(Workbook workbook, CellAddress cellAddress) {
		List<String> customerList = new ArrayList();
		int numOfSheets = workbook.getNumberOfSheets();
		Sheet hssfSheet;
		int columnIndex = cellAddress.getColumn();
		for (int i = 0; i < numOfSheets; i++) {
			hssfSheet = (Sheet) workbook.getSheetAt(i);
			customerList.addAll(getCustomerNamesFromSheet(hssfSheet, columnIndex));
		}
		return new HashSet<String>(customerList);
	}

	//gets customer names from one sheet
	public static List<String> getCustomerNamesFromSheet(Sheet hssfSheet, int columnIndex) {
		List<String> customerNames = new ArrayList<String>();
		Iterator<Row> rowIterator = hssfSheet.iterator();
		Iterator<Cell> cellIterator;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				//checking if in proper column and in values section
				if ((cell.getColumnIndex() == columnIndex) && (cell.getRowIndex() != 0)) {
					//remove un-needed characters from name
					String cellVal = cell.getStringCellValue();
					if (!cellVal.equals(cellVal.replaceAll("[\\-\\+\\.\\^:,/]", ""))) {
						System.out.println("old Cell Value: " + cellVal + " changed to new cell value: " + cellVal.replaceAll("[\\-\\+\\.\\^:,/]", ""));
						cell.setCellValue(cell.getStringCellValue().replaceAll("[\\-\\+\\.\\^:,/]", ""));
					}
					customerNames.add(cell.getStringCellValue());
					break;
				}
			}
		}
		return customerNames;
	}

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

	//copies file
	public static void copyFile(String srcPath, String desPath) {
		try {
			FileUtils.copyFile(new File(srcPath), new File(desPath));
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("copied file: " + srcPath + " successfully");
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

	//loads workbook
	public static Workbook loadWorkbook(String path) {
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

	//read customer ids file, returns ID to Name map
	public static Map<String, Integer> loadRegionToCountryMap(Sheet sheet, int idColumn, int nameColumn) {
		Row row;
		Map<String, Integer> regionToCountryMap = new HashMap<String, Integer>();
		DataFormatter dataFormatter = new DataFormatter();

		//iterate over all rows and cells in workbook
		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			String countryName = dataFormatter.formatCellValue(row.getCell(nameColumn));
			String regionId = dataFormatter.formatCellValue(row.getCell(idColumn));
			regionToCountryMap.put(countryName, Integer.parseInt(regionId));
		}
		return regionToCountryMap;
	}

	//calculates frieght for all customers
	public static void calcFreightForAllCustomers(Map<Customer, Workbook> customersMap, Map<String, Integer> regionsMap) {
		for (Customer customer : customersMap.keySet()) {
			System.out.println("started calcCustomerFreight on : " + customer.getName());
			//load customer price list
			final Workbook custPriceListWb = loadWorkbook("inputdir/customer price lists/" + customer.getName() + " " + Constants.PL_FILE_ENDING);
			calcCustomerFreight(customersMap.get(customer), custPriceListWb, regionsMap);
		}
	}

	/*
	calculates freight cell value according to formula
	CHARGE = [(customer price per region) X weight]  + [(fuel surcharge% * customer price per region)]
*/

	public static void calcCustomerFreight(Workbook workbook, Workbook custPriceListWb, Map<String, Integer> regionsMap) {
		final Sheet sheet = workbook.getSheet(Constants.SHEET_NAME);
		int weightCol = -1;
		int destinationCol = -1;
		int freightCol = -1;

		try {
			weightCol = findCellByName(sheet, Constants.WEIGHT_COL).getColumn();
			destinationCol = findCellByName(sheet, Constants.DES_COL).getColumn();
			freightCol = findCellByName(sheet, Constants.FREIGHT).getColumn();
		} catch (Exception e) {
			e.printStackTrace();
		}

		double fuelScp = getFuelSurchargePercent();
		if(fuelScp < 0 || fuelScp > 1){
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
		//iterate over all rows in workbook
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			//get current row
			row = sheet.getRow(i);
			//get weight value
			cell = row.getCell(weightCol);
			weight = Double.parseDouble(dataFormatter.formatCellValue(cell));
			//get shipment des country
			cell = row.getCell(destinationCol);
			country = dataFormatter.formatCellValue(cell);

			//get zone for country
			zone = regionsMap.get(country).intValue();
			/*
			//started calcCustomerFreight on : GOODU ART LTD
			java.lang.NullPointerException
			at Controller.SmartShipApplication.calcCustomerFreight(SmartShipApplication.java:739)
			 */

			//calc price according to weight and zone
			pricePerWeightAndZone = getCustomerPrice(custPriceListWb, weight, zone);
			totalPrice = ( (1+ fuelScp) * pricePerWeightAndZone ); 					//fuelScp range:[0 - 1]

			//write updated price value to cell
			cell = row.getCell(freightCol);
			cell.setCellValue(totalPrice);
			System.out.println("updated cell : " + cell.getAddress() + " with a value of: " + totalPrice);
		}
	}

	/**
	 * gets the price per zone for a given customer
	 * //find closest base weight and get price example: weight is 11.5 , base =11K  D5-d53
	 * Zone Cells: E4-K4, values: [1-7], ZONE_OFFSET = 3
	 */
	public static double getCustomerPrice(Workbook custPriceListWb, double weight, int zone) {
		Sheet sheet = custPriceListWb.getSheet(Constants.SHEET_NAME); //TODO change name to constant in english
		Row row;
		Cell cell;
		double baseWeight = 0;
		double basePrice = 0;
		double diff;
		double additionalPrice = 0;
		final int zoneCol = zone + Constants.ZONE_OFFSET;
		final int startRow = 4;
		final int endRow = 52;
		final int baseWeightCol = 3;

		//iterate over rows
		for (int i = startRow; i <= endRow; i++) {
			row = sheet.getRow(i);
			cell = row.getCell(baseWeightCol);
			//not to overflow
			if (i + 1 <= endRow) {
				//find closest weight value
				if ((weight >= cell.getNumericCellValue()) && (weight < sheet.getRow(i + 1).getCell(baseWeightCol).getNumericCellValue())) {
					baseWeight = cell.getNumericCellValue();
					basePrice = (baseWeight * (sheet.getRow(i).getCell(zoneCol).getNumericCellValue()));
					break;
				}
			}
			//if last row get last value
			else {
				baseWeight = cell.getNumericCellValue();
			}
		}

		//calc diff between table base weight and actual weight
		diff = weight - baseWeight;

		//add additional weight according to addition table, all cases:
		if (weight >= 10 && weight < 20) {
			additionalPrice = ((diff / Constants.WEIGHT_MULTIPLIER) * (sheet.getRow(57).getCell(zoneCol).getNumericCellValue()));
		} else if (weight >= 20 && weight < 30) {
			additionalPrice = ((diff / Constants.WEIGHT_MULTIPLIER) * (sheet.getRow(58).getCell(zoneCol).getNumericCellValue()));
		} else if (weight >= 30 && weight < 70) {
			additionalPrice = ((diff / Constants.WEIGHT_MULTIPLIER) * (sheet.getRow(59).getCell(zoneCol).getNumericCellValue()));
		} else if (weight >= 70 && weight < 300) {
			additionalPrice = ((diff / Constants.WEIGHT_MULTIPLIER) * (sheet.getRow(60).getCell(zoneCol).getNumericCellValue()));
		} else if (weight >= 300){
			additionalPrice = ((diff / Constants.WEIGHT_MULTIPLIER) * (sheet.getRow(61).getCell(zoneCol).getNumericCellValue()));
		}

		return basePrice + additionalPrice;
	}


	public static double getFuelSurchargePercent() {
		//TODO
		//get getFuelSurchargePercent from somewhere
		return 0;
	}


}

/*
		utilities:
		//create files price list per customer
			for(String s: customerIdToNameMap.values()){
				copyFile("inputdir/customer price lists/bananot price list.xlsx","inputdir/customer price lists/" + s + " " + Constants.PL_FILE_ENDING);
			}
 */