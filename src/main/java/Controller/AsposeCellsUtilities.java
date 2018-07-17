package Controller;

import Model.Constants;
import com.aspose.cells.Cells;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.Map;


public class AsposeCellsUtilities {

    public static void main(String[] args) {
        try {
            Workbook workbook = new Workbook("outputdir/customers/MARCAS BRANDS.xls");
            //Worksheet worksheet = workbook.getWorksheets().get(0);
            //Cells cells = worksheet.getCells();
            //cells.deleteBlankColumns();
            //worksheet.autoFitColumns();
            //Save the document in PDF format
            workbook.save("outputdir/customers/" + "AAAAAAA" + System.currentTimeMillis() + " AsposeConvert.pdf", SaveFormat.PDF);
            //workbook.save("outputdir/customers/" + System.currentTimeMillis() + " AsposeConvert.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //load wb files, delete empty columns, auto-fits columns saves and closes WB files
    public static void deleteBlankAndAutoFitColumns(Map<String, org.apache.poi.ss.usermodel.Workbook> mapCustomerFileNameToWb) {
        org.apache.poi.ss.usermodel.Workbook wb;
        try {
            for (String customerFileName : mapCustomerFileNameToWb.keySet()) {
                Workbook workbook = new Workbook("outputdir/customers/" + customerFileName + ".xls");
                Worksheet worksheet = workbook.getWorksheets().get(Constants.FIRST_SHEET_NUM);
                Cells cells = worksheet.getCells();
                cells.deleteBlankColumns();
                worksheet.autoFitColumns();
                workbook.save("outputdir/customers/" + customerFileName + ".xls");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
