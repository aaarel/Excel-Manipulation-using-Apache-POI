package Controller;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class TestRunner {

    public static void main(String[] args) {
        //test
        System.out.println("******************************");
        System.out.println("******************************");
        System.out.println("******************************");
        System.out.println("******************************");
        System.out.println("******************************");

        Workbook wb1 = UtilityMethods.loadWb("outputdir/customers/MARCAS BRANDS.xls");
        Sheet sheet1 = wb1.getSheetAt(0);
        //UtilityMethods.deleteColumn(sheet1, UtilityMethods.findCellByName(sheet1, "Acnt").getColumn());
        Map<String, Workbook> mapCustomerFileNameWb1 = new HashMap<String, Workbook>();
        mapCustomerFileNameWb1.put("MARCAS BRANDS.xls", wb1);
        //UtilityMethods.cleanColumnsFromCustomerWbFiles2(mapCustomerFileNameWb1);
        //UtilityMethods.saveAndCloseWbFiles(mapCustomerFileNameWb1);

    }
}
