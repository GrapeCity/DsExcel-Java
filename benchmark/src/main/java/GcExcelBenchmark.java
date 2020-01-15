import com.grapecity.documents.excel.*;

import java.util.Date;
import java.util.Random;

public class GcExcelBenchmark {

    public static void TestSetRangeValues_Double(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        System.out.println();
        System.out.println(String.format("GcExcel benchmark for double values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        IWorkbook workbook = new Workbook();
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        long start = System.currentTimeMillis();

        double[][] values = new double[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] = i + j;
            }
        }
        worksheet.getRange(0, 0, rowCount, columnCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel set double values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getRange(0, 0, rowCount, columnCount).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel get double values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/gcexcel-saved-doubles.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel save doubles to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem.value));
    }


    public static void TestSetRangeValues_String(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        System.out.println();
        System.out.println(String.format("GcExcel benchmark for string values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        IWorkbook workbook = new Workbook();
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        Random random = new Random();
        String AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        long start = System.currentTimeMillis();

        String[][] values = new String[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] =  String.valueOf(AlphaNumericString.charAt(random.nextInt(25)));
            }
        }

        worksheet.getRange(0, 0, rowCount, columnCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel set string values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getRange(0, 0, rowCount, columnCount).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel get string values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/gcexcel-saved-string.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel save string to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem.value));
    }

    public static void TestSetRangeValues_Date(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        System.out.println();
        System.out.println(String.format("GcExcel benchmark for date values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        IWorkbook workbook = new Workbook();
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        long start = System.currentTimeMillis();

        Date[][] values = new Date[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] = new Date();
            }
        }
        worksheet.getRange(0, 0, rowCount, columnCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel set date values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getRange(0, 0, rowCount, columnCount).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel get date values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/gcexcel-saved-date.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel save date to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem.value));
    }

    public static void TestSetRangeFormulas(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> calcTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        System.out.println();
        System.out.println(String.format("GcExcel benchmark for double values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        IWorkbook workbook = new Workbook();
        workbook.setReferenceStyle(ReferenceStyle.R1C1);
        IWorksheet worksheet = workbook.getWorksheets().get(0);


        double[][] values = new double[rowCount][2];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < 2; j++) {
                values[i][j] = i + j;
            }
        }
        worksheet.getRange(0, 0, rowCount, 2).setValue(values);

        long start = System.currentTimeMillis();
        worksheet.getRange(0, 2, rowCount, columnCount).setFormula("=SUM(RC[-2],RC[-1])");
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel set formulas: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        workbook.calculate();
        end = System.currentTimeMillis();

        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel calculates formula: %.1f s", calcTime.value));

        workbook.setReferenceStyle(ReferenceStyle.A1);

        start = System.currentTimeMillis();
        workbook.save("output/gcexcel-saved-formulas.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel save formulas to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem.value));
    }


    public static void TestBigExcelFile(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> openTime,
                                                 RefObject<Double> calcTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        System.out.println();
        System.out.println(String.format("GcExcel benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        IWorkbook workbook = new Workbook();

        long start = System.currentTimeMillis();
        workbook.open("files/test-performance.xlsx");
        long end = System.currentTimeMillis();

        openTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel open big Excel: %.1f s", openTime.value));

        start = System.currentTimeMillis();
        workbook.dirty();
        workbook.calculate();
        end = System.currentTimeMillis();

        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel calculate formulas for big Excel: %.1f s", calcTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/gcexcel-saved-test-performance.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("GcExcel save back to big Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("GcExcel used memory: %.1f MB", usedMem.value));
    }
}
