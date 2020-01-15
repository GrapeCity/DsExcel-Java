
import com.grapecity.documents.excel.CellInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Random;

import java.util.Random;

public class POIBenchmark {

    public static void TestSetRangeValues_Double(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("POI benchmark for double values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet worksheet = workbook.createSheet("poi");

        Random rand = new Random();
        long start = System.currentTimeMillis();

        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.createRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.createCell(c).setCellValue(rand.nextDouble());
            }
        }
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI set double values: %.1fs", setTime.value));

        start = System.currentTimeMillis();
        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.getRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.getCell(c).getNumericCellValue();
            }
        }
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI get double values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("output/poi-saved-doubles.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI save doubles to Excel: %.1f s", saveTime.value));


        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("POI used memory: %.1f MB", usedMem.value));
    }


    public static void TestSetRangeValues_String(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("POI benchmark for string values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet worksheet = workbook.createSheet("poi");


        Random random = new Random();
        String AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        long start = System.currentTimeMillis();

        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.createRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.createCell(c).setCellValue(String.valueOf(AlphaNumericString.charAt(random.nextInt(25))));
            }
        }
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI set string values: %.1fs", setTime.value));

        start = System.currentTimeMillis();
        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.getRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.getCell(c).getStringCellValue();
            }
        }
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI get string values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("output/poi-saved-string.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI save string to Excel: %.1f s", saveTime.value));


        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("POI used memory: %.1f MB", usedMem.value));
    }

    public static void TestSetRangeValues_Date(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("POI benchmark for date values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet worksheet = workbook.createSheet("poi");

        long start = System.currentTimeMillis();

        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.createRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.createCell(c).setCellValue(new Date());
            }
        }
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI set date values: %.1fs", setTime.value));

        start = System.currentTimeMillis();
        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.getRow(r);
            for (int c = 0; c < columnCount; c++) {
                row.getCell(c).getDateCellValue();
            }
        }
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI get date values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("output/poi-saved-doubles.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI save date to Excel: %.1f s", saveTime.value));


        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//        System.out.println(String.format("POI used memory: %.1f MB", usedMem.value));
    }

    public static void TestSetRangeFormulas(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> calcTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("POI benchmark for formulas values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet worksheet = workbook.createSheet("poi");

        Random rand = new Random();

        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.createRow(r);
            for (int c = 0; c < 2; c++) {
                row.createCell(c).setCellValue(r + c);
            }
        }

        long start = System.currentTimeMillis();
        for (int r = 0; r < rowCount; r++) {
            Row row = worksheet.createRow(r);
            for (int c = 2; c < columnCount; c++) {
                Cell cell = row.createCell(c);
                CellReference reference1 = new CellReference(r, c - 2);
                CellReference reference2 = new CellReference(r, c - 1);
                cell.setCellFormula(String.format("SUM(%s, %s)", reference1.formatAsString(), reference2.formatAsString()));
            }
        }
        long end = System.currentTimeMillis();
        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI set formulas values: %.1fs", setTime.value));

        start = System.currentTimeMillis();
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
        end = System.currentTimeMillis();

        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI calculate formulas: %.1f s", calcTime.value));

        start = System.currentTimeMillis();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("output/poi-saved-formulas.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI save formulas to Excel: %.1f s", saveTime.value));


        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//        System.out.println(String.format("POI used memory: %.1f MB", usedMem.value));
    }

    public static void TestBigExcelFile(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> openTime,
                                                 RefObject<Double> calcTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("POI benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();


        long start = System.currentTimeMillis();
        XSSFWorkbook workbook = new XSSFWorkbook("files/test-performance.xlsx");
        long end = System.currentTimeMillis();

        openTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI open big Excel: %.1fs", openTime.value));

        start = System.currentTimeMillis();
        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
        calcTime.value = (end - start) * 0.001;
        end = System.currentTimeMillis();
        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI calculate formulas for big Excel: %.1f s", calcTime.value));

        start = System.currentTimeMillis();
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("output/poi-saved-test-performance.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("POI save back to big Excel: %.1f s", saveTime.value));


        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("POI used memory: %.1f MB", usedMem.value));
    }
}
