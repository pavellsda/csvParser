import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
/**
 * Created by pavellsda on 08.02.17.
 */
public class Base {

    private Utils utils = new Utils();

    void getProducts(Order[] orders) throws IOException, InvalidFormatException {

        Map<String,String> warhammerOrders = new HashMap<String,String>();
        Map<String,String> siteOrders = new HashMap<String,String>();

        String warhammerOrdersFile = "inputData/gwof.xlsx";
        String siteOrdersFile = "inputData/products.xlsx";

        DataFormatter formatter = new DataFormatter(Locale.US);

        Workbook warhammerOrdersBook = new XSSFWorkbook(new FileInputStream(warhammerOrdersFile));
        Sheet warhammerOrdersSheet = warhammerOrdersBook.getSheet("NE Trade");


        for(int i = 0; i < warhammerOrdersSheet.getLastRowNum(); i++) {
            Row row = warhammerOrdersSheet.getRow(i);
            String attribute = null;
            String name = null;


            if (row.getCell(8).getCellType() == Cell.CELL_TYPE_STRING) {
                name = row.getCell(8).getStringCellValue();
            }
            if (row.getCell(7).getCellType() == Cell.CELL_TYPE_STRING) {
                attribute = row.getCell(7).getStringCellValue();
            }

            if(name!=null&&attribute!=null)
                warhammerOrders.put(attribute, name);

        }

        warhammerOrdersBook.close();

        Workbook siteOrdersBook = new XSSFWorkbook(new FileInputStream(siteOrdersFile));
        Sheet siteOrdersSheet = siteOrdersBook.getSheet("Sheet1");

        for(int i = 0; i < siteOrdersSheet.getLastRowNum(); i++) {
            Row row = siteOrdersSheet.getRow(i);
            String attribute = formatter.formatCellValue(row.getCell(1));
            String name = formatter.formatCellValue(row.getCell(0));
            if(name!=null&&attribute!=null){
                siteOrders.put(name, attribute);
            }

        }
        siteOrdersBook.close();
        Thread thr = new Thread(() -> {
            try {
                System.out.println("\nСоздание таблиц excel.\n");
                utils.createExcelFiles(orders, siteOrders, warhammerOrders);
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        thr.start();


        warhammerOrdersBook = new XSSFWorkbook(new FileInputStream(warhammerOrdersFile));
        warhammerOrdersSheet = warhammerOrdersBook.getSheet("NE Trade");

        for(int i = 0; i < warhammerOrdersSheet.getLastRowNum(); i++) {
            Row row = warhammerOrdersSheet.getRow(i);
            String attribute = null;
            String name = null;

            if (row.getCell(7).getCellType() == Cell.CELL_TYPE_STRING) {
                attribute = row.getCell(7).getStringCellValue();
                for (Order order : orders) {
                    if (Objects.equals(attribute, order.getAttribute())) {
                        Cell cell = row.getCell(9);
                        cell.setCellValue(order.getCount());
                    }
                }
            }

        }
        File gfow = new File(CsvParser.getPathToDir() + "/gwof" +
                CsvParser.getDate() + ".xlsx");
        FileOutputStream fileOut = new FileOutputStream(gfow);
        System.out.println("Таблица gfow.xlsx создана и находится по адресу: \n" +
                gfow.getAbsolutePath());
        warhammerOrdersBook.write(fileOut);
        fileOut.close();
        warhammerOrdersBook.close();



    }
}
