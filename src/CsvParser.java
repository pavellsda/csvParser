import com.opencsv.CSVReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by pavellsda on 07.02.17.
 */
public class CsvParser {

    private static double  summSubTotal = 0;
    private static double  summTotal    = 0;
    private static double  summOrders   = 0;
    private static int     ordCount     = 0;
    private static boolean success      = false;
    private static String  pathToDir    = null;
    private static String  date         = null;

    private static Vector<String> productsList;
    private static List<Order>    ordersFromCSV;
    private static List<Order>    ordersList;


    public static void main(String[] argv){

        createDir("results/" + getDirDate());

        String fileToParse = "inputData/Orders.csv";

        ordersList = new ArrayList<>();
        ordersFromCSV = new ArrayList<>();

        print("Загружается список всех товаров в магазине.\n");
        productsList = getProductsList();

        print("Список всех товаров загружен.\n" +
                "Загружается список заказов из Orders.csv.\n");
        readCSV(fileToParse);

        try {
            System.out.println("\nСравнение артикулов.");
           new Base().getProducts(ordersList.toArray(new Order[ordersList.size()]));
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }


    }

    private static void print(String message){
        System.out.println(message);
    }

    private static void readCSV(String fileToParse){
        try {
            CSVReader csvReader = new CSVReader(new FileReader(fileToParse));
            List content = csvReader.readAll();

            print("Список заказов загружен.\n" +
                    "Происходит разбор списка заказов.\n");

            int procent = 0;
            for (int i = 0; i < content.size(); i++) {
                parseCSV(content.get(i));
                int tempProcent = i * 100 / (content.size() - 1);
                if(tempProcent!=procent) {
                    procent = tempProcent;
                    print("Выполнено: " + procent + "%");
                }
            }

            csvReader.close();

            String[] ordersWithoutD = new String[ordersFromCSV.size()];
            for(int i = 0; i < ordersFromCSV.size(); i++) {
                ordersWithoutD[i] = ordersFromCSV.get(i).getName();
            }

            Set<String> set = new HashSet<>(Arrays.asList(ordersWithoutD));
            String[] result = set.toArray(new String[set.size()]);

            for (String aResult : result) {
                ordersList.add(new Order(aResult, 0));
            }

            for (Order anOrdersFromCSV : ordersFromCSV) {
                for (Order anOrdersList : ordersList) {
                    if (Objects.equals(anOrdersFromCSV.getName(), anOrdersList.getName())) {
                        anOrdersList.incCount(anOrdersFromCSV.getCount());
                    }
                }

            }

        }
        catch (Exception e) {
           e.printStackTrace();
        }
    }

    private static void parseCSV(Object object){
        String[] row = (String[]) object;
        for (String atRow : row) {
            atRow = atRow.replaceAll("\"","");

            if (Objects.equals(atRow, "Выполнен"))
                success = false;
            if (Objects.equals(atRow, "Отменен"))
                success = false;
            if (Objects.equals(atRow, "На удержании"))
                success = false;
            if (Objects.equals(atRow, "Ожидает оплаты"))
                success = false;
            if (Objects.equals(atRow, "Оплаченный предзаказ"))
                success = true;
        }

        for (int i = 0; i < row.length; i++) {

            if (success) {
                if (Objects.equals(row[i], "&quot;Iguana&quot; Squadron")) {
                    row[i] = "\"Iguana\" Squadron";
                }
                if (Objects.equals(row[i], "&#8216;Ardcoat")) {
                    row[i] = "'Ardcoat";
                }

                for (String aProductsList : productsList) {
                    if (Objects.equals(row[i], aProductsList)) {
                        ordersFromCSV.add(new Order(row[i], Integer.parseInt(row[i + 1])));
                        summOrders += Double.parseDouble(row[i + 2]);
                    }
                }

                if (Objects.equals(row[i], "sub_total")) {
                    summSubTotal += Double.parseDouble(row[i + 2]);
                }
                if (Objects.equals(row[i], "total")) {
                    summTotal += Double.parseDouble(row[i + 2]);
                    ordCount++;
                }
            }

        }
    }

    static void printResults(){
        System.out.println("\nПроисходит создание отладочного файла и отчета.\n");
        Debug(ordersList);

        System.out.println("\nВсего заказов: " + ordCount);
        System.out.println("Сумма с доставкой: " + summTotal);
        System.out.println("Сумма без учета доставки: " + summSubTotal);
        System.out.println("Сумма(если считать по товарам): " + summOrders);

    }

    private static void Debug(List<Order> orderList){
        int countOrders = 0;

        try (FileWriter writer = new FileWriter(getPathToDir() + "/output" +
                getDate() + ".txt", false)) {
            for (int i = 1; i < orderList.size(); i++) {
                writer.write(orderList.get(i).getName() + "|" + orderList.get(i).getCount());
                countOrders = countOrders + orderList.get(i).getCount();
                writer.append('\n');

                writer.flush();
            }

        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }


        try (FileWriter writer = new FileWriter(getPathToDir() + "/result" +
                getDateStr() + ".txt", false)) {
            writer.write("Всего заказов: " + ordCount);
            writer.append('\n');

            writer.write("Всего Товаров: " + countOrders);
            writer.append('\n');

            writer.write("Сумма с доставкой: " + summTotal);
            writer.append('\n');

            writer.write("Сумма без учета доставки: " + summSubTotal);
            writer.append('\n');

            writer.write("Сумма(если считать по товарам): " + summOrders);
            writer.append('\n');

            writer.flush();

        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }
        System.out.println("Отладочный файл создан.(output.txt)\n");
        System.out.println("Отчет создан.(result.txt)\n");
    }
    private static Vector<String> getProductsList() {
        String siteOrdersFile = "inputData/products.xlsx";
        Vector<String> productsList = new Vector<>();

        DataFormatter formatter = new DataFormatter(Locale.US);

        Workbook siteOrdersBook;
        try {
            siteOrdersBook = new XSSFWorkbook(new FileInputStream(siteOrdersFile));
            Sheet siteOrdersSheet = siteOrdersBook.getSheet("Sheet1");

            for(int i = 0; i < siteOrdersSheet.getLastRowNum(); i++) {
                Row row = siteOrdersSheet.getRow(i);
                String name = formatter.formatCellValue(row.getCell(0));
                productsList.add(name);

            }
            siteOrdersBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return productsList;
    }

    private static boolean createDir(String path){
        pathToDir = path;
        return new File(path).mkdir();
    }

    private static String getDirDate(){
        long curTime = System.currentTimeMillis();

        return new SimpleDateFormat("dd.MM.yyyy|hh.mm.ss").format(curTime);
    }
    static String getDate(){

        long curTime = System.currentTimeMillis();

        date = new SimpleDateFormat("_dd.MM.yyyy").format(curTime);

        return new SimpleDateFormat("_dd.MM.yyyy").format(curTime);
    }

    static String getPathToDir() {
        return pathToDir;
    }

    public static String getDateStr() {
         return date;
    }
}

