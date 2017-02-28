import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by pavellsda on 09.02.17.
 */
public class Utils {
    private CellStyle getCellStyle(Workbook book, String font, String alin){
        CellStyle style = book.createCellStyle();
        switch (font) {
            case "YELLOW":
                style.setFillForegroundColor(HSSFColor.YELLOW.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case "RED":
                style.setFillForegroundColor(HSSFColor.RED.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case "GREY":
                style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case "GREEN":
                style.setFillForegroundColor(HSSFColor.GREEN.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case "BLUE":
                style.setFillForegroundColor(HSSFColor.BLUE.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case "WHITE":
                style.setFillForegroundColor(HSSFColor.WHITE.index);
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
        }
        switch (alin){
            case "LEFT":
                style.setAlignment(HorizontalAlignment.LEFT);
                break;
            case "CENTER":
                style.setAlignment(HorizontalAlignment.CENTER);
                break;
            case "RIGHT":
                style.setAlignment(HorizontalAlignment.RIGHT);
                break;
        }
        style.setBorderLeft(BorderStyle.HAIR);
        style.setBorderRight(BorderStyle.HAIR);
        style.setBorderBottom(BorderStyle.HAIR);
        style.setBorderTop(BorderStyle.HAIR);

        return style;
    }

    private static Sheet deleteSpace(Sheet sheet){
        boolean isRowEmpty = false;
        for(int i = 0; i < sheet.getLastRowNum(); i++){
            if(sheet.getRow(i)==null){
                isRowEmpty=true;
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                i--;
                continue;
            }
            for(int j =0; j<sheet.getRow(i).getLastCellNum();j++){
                if(sheet.getRow(i).getCell(j).toString().trim().equals("")){
                    isRowEmpty=true;
                }else {
                    isRowEmpty=false;
                    break;
                }
            }
            if(isRowEmpty){
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
                i--;
            }
        }
        return sheet;
    }


    void createExcelFiles(Order[] orders, Map<String, String> siteOrders,
                          Map<String, String> warhammerOrders) throws IOException {

        Workbook warhammerBook = new XSSFWorkbook();
        Sheet warhammerSheet = warhammerBook.createSheet("ORDERS");

        Workbook notWarhammerBook = new XSSFWorkbook();
        Sheet notWarhammerSheet = notWarhammerBook.createSheet("ORDERS");

        for (Order order : orders) {
            order.setAttribute(siteOrders.get(order.getName()));
            order.setSystemName(warhammerOrders.get(order.getAttribute()));

            if (order.getSystemName() != null) {
                if (order.getSystemName().contains("6-PACK")) {
                    double count = order.getCount() / 6;
                    int result;
                    if (order.getCount() % 6 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }

                if (order.getSystemName().contains("10-PACK")) {
                    double count = order.getCount() / 10;
                    int result;
                    if (order.getCount() % 10 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }

                if (order.getSystemName().contains("x10")) {
                    double count = order.getCount() / 10;
                    int result;
                    if (order.getCount() % 10 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }

                if (order.getSystemName().contains("50-PACK")) {
                    double count = order.getCount() / 50;
                    int result;
                    if (order.getCount() % 50 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }
                if (order.getSystemName().contains("3-PACK")) {
                    double count = order.getCount() / 3;
                    int result;
                    if (order.getCount() % 3 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }

                if (order.getSystemName().contains("PACK OF 3")) {
                    double count = order.getCount() / 3;
                    int result;
                    if (order.getCount() % 3 == 0)
                        result = (int) count;
                    else
                        result = (int) count + 1;
                    order.setCount(result);
                }
            }

        }

        Row title = warhammerSheet.createRow(0);
        Row titleNW = notWarhammerSheet.createRow(0);

        Cell systemTitle = title.createCell(0);
        systemTitle.setCellStyle(getCellStyle(warhammerBook, "GREY", "CENTER"));
        systemTitle.setCellValue("Product name on the site");

        Cell siteTitle = title.createCell(1);
        siteTitle.setCellStyle(getCellStyle(warhammerBook, "GREY", "CENTER"));
        siteTitle.setCellValue("Name");

        Cell attributeTitle = title.createCell(2);
        attributeTitle.setCellStyle(getCellStyle(warhammerBook, "GREY", "CENTER"));
        attributeTitle.setCellValue("Vendor code");

        Cell countTitle = title.createCell(3);
        countTitle.setCellStyle(getCellStyle(warhammerBook, "GREY", "CENTER"));
        countTitle.setCellValue("Order");

        Cell siteTitleNW = titleNW.createCell(0);
        siteTitleNW.setCellStyle(getCellStyle(notWarhammerBook, "GREY", "CENTER"));
        siteTitleNW.setCellValue("Name");

        Cell attributeTitleNW = titleNW.createCell(1);
        attributeTitleNW.setCellStyle(getCellStyle(notWarhammerBook, "GREY", "CENTER"));
        attributeTitleNW.setCellValue("Vendor code");

        Cell countTitleNW = titleNW.createCell(2);
        countTitleNW.setCellStyle(getCellStyle(notWarhammerBook, "GREY", "CENTER"));
        countTitleNW.setCellValue("Order");

        List<String> list = new ArrayList<String>();

        for(int i = 2; i <= orders.length; i++){

            if(warhammerOrders.get(orders[i-1].getAttribute())==null){
                if(orders[i-1].getAttribute()!=null) {

                    list.add(orders[i-1].getAttribute());
                }
                else {
                    Row row = notWarhammerSheet.createRow(i-1);

                    Cell name = row.createCell(0);
                    name.setCellValue(orders[i-1].getName());

                    Cell attr = row.createCell(1);
                    attr.setCellStyle(getCellStyle(notWarhammerBook,"WHITE","CENTER"));
                    attr.setCellValue(orders[i-1].getAttribute());

                    Cell count = row.createCell(2);
                    count.setCellStyle(getCellStyle(notWarhammerBook,"YELLOW","CENTER"));
                    count.setCellValue(orders[i-1].getCount());
                }

            }
            else{

                Row row = warhammerSheet.createRow(i);

                Cell sysName = row.createCell(0);
                sysName.setCellValue(orders[i-1].getSystemName());

                Cell name = row.createCell(1);
                name.setCellValue(orders[i-1].getName());

                Cell attr = row.createCell(2);
                attr.setCellStyle(getCellStyle(warhammerBook,"WHITE","CENTER"));
                attr.setCellValue(orders[i-1].getAttribute());

                Cell count = row.createCell(3);
                count.setCellStyle(getCellStyle(warhammerBook,"YELLOW","CENTER"));
                count.setCellValue(orders[i-1].getCount());


            }

        }

        deleteSpace(notWarhammerSheet);

        list.sort((o1, o2) -> (o1.length() - o2.length()));

        for(int i = 0; i < list.size(); i++){
                Row row = notWarhammerSheet.createRow(notWarhammerSheet.getLastRowNum()+i);

                for(int j = 0; j < orders.length; j++){
                    if(Objects.equals(orders[j].getAttribute(), list.get(i))){
                        Cell name = row.createCell(0);
                        name.setCellValue(orders[j].getName());

                        Cell attr = row.createCell(1);
                        attr.setCellStyle(getCellStyle(notWarhammerBook,"WHITE","CENTER"));
                        attr.setCellValue(orders[j].getAttribute());

                        Cell count = row.createCell(2);
                        count.setCellStyle(getCellStyle(notWarhammerBook,"YELLOW","CENTER"));
                        count.setCellValue(orders[j].getCount());
                    }
                }


        }

        deleteSpace(warhammerSheet);
        deleteSpace(notWarhammerSheet);

        resize(warhammerBook);
        resize(notWarhammerBook);

        File wh = new File(CsvParser.getPathToDir() + "/WH" +
                CsvParser.getDateStr() + ".xlsx");
        warhammerBook.write(new FileOutputStream(wh));
        warhammerBook.close();

        File notWh = new File(CsvParser.getPathToDir() + "/NotWH" +
                CsvParser.getDateStr() + ".xlsx");
        notWarhammerBook.write(new FileOutputStream(notWh));
        notWarhammerBook.close();

        System.out.println("\nТаблицы WH.xlsx и NotWH.xlsx созданы и находятся по адресу: \n" +
                "WH.xlsx - " + wh.getAbsolutePath()+"\n" +
                "NotWH.xlsx - " + notWh.getAbsolutePath());

        CsvParser.printResults();
    }

    private void resize(Workbook book){
        Row row = book.getSheetAt(0).getRow(0);
        for(int colNum = 0; colNum<row.getLastCellNum();colNum++)
            book.getSheetAt(0).autoSizeColumn(colNum);
    }



}
