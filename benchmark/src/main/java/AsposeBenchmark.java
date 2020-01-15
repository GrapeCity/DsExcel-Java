
import com.aspose.cells.*;
import com.grapecity.documents.excel.CellInfo;

import java.util.Date;
import java.util.Random;

public class AsposeBenchmark {

    public static void TestSetRangeValues_Double(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("Aspose.Cells benchmark for double values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        long start = System.currentTimeMillis();

        Double[][] values = new Double[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] = (double)(i + j);
            }
        }

        worksheet.getCells().createRange("A1:AC" + rowCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells set double values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getCells().createRange("A1:AC" + String.valueOf(rowCount)).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells get double values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/aspose-saved-doubles.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells save doubles to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("Aspose.Cells used memory: %.1f MB", usedMem.value));

    }

    public static void TestSetRangeValues_String(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("Aspose.Cells benchmark for string values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Random random = new Random();
        String AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        long start = System.currentTimeMillis();

        String[][] values = new String[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] = String.valueOf(AlphaNumericString.charAt(random.nextInt(25)));
            }
        }

        worksheet.getCells().createRange("A1:AC" + rowCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells set string values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getCells().createRange("A1:AC" + String.valueOf(rowCount)).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells get string values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/aspose-saved-string.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells save string to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("Aspose.Cells used memory: %.1f MB", usedMem.value));

    }

    public static void TestSetRangeValues_Date(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("Aspose.Cells benchmark for date values with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        long start = System.currentTimeMillis();

        Date[][] values = new Date[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                values[i][j] = new Date();
            }
        }

        worksheet.getCells().createRange("A1:AC" + rowCount).setValue(values);
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells set date values: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        Object tmpValues = worksheet.getCells().createRange("A1:AC" + String.valueOf(rowCount)).getValue();
        end = System.currentTimeMillis();

        getTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells get date values: %.1f s", getTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/aspose-saved-date.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells save date to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//        System.out.println(String.format("Aspose.Cells used memory: %.1f MB", usedMem.value));

    }

    public static void TestSetRangeFormulas(int rowCount,
                                                 int columnCount,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> calcTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("Aspose.Cells benchmark for formulas with %,d rows and %d columns", rowCount, columnCount));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        Double[][] values = new Double[rowCount][columnCount];

        for (int i = 0; i < rowCount; i++) {
            for (int j = 0; j < 2; j++) {
                values[i][j] = (double)(i + j);
            }
        }
        worksheet.getCells().createRange("A1:B" + rowCount).setValue(values);

        long start = System.currentTimeMillis();
        Cells cells = worksheet.getCells();
        for(int c = 2; c < columnCount; c++){
            cells.getCell(0, c).setSharedFormula(
                String.format("SUM(%s1, %s1)", CellInfo.ColumnIndexToName(c - 2), CellInfo.ColumnIndexToName(c -1)),
                rowCount, c);
        }
        long end = System.currentTimeMillis();

        setTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells set formulas: %.1f s", setTime.value));

        start = System.currentTimeMillis();
        workbook.calculateFormula();
        end = System.currentTimeMillis();

        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells calculate formulas: %.1f s", calcTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/aspose-saved-formulas.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells save formulas to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
//      System.out.println(String.format("Aspose.Cells used memory: %.1f MB", usedMem.value));

    }

    public static void TestBigExcelFile(int rowCount,
                                            int columnCount,
                                            RefObject<Double> openTime,
                                            RefObject<Double> calcTime,
                                            RefObject<Double> saveTime,
                                            RefObject<Double> usedMem) throws Exception {
        System.out.println();
        System.out.println(String.format("Aspose.Cells benchmark for test-performance.xlsx which is 20.5MB with a lot of values, formulas and styles"));

        double startMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();


        long start = System.currentTimeMillis();
        Workbook workbook = new Workbook("files/test-performance.xlsx");
        long end = System.currentTimeMillis();

        openTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells open big Excel file: %.1f s", openTime.value));

        start = System.currentTimeMillis();
        workbook.calculateFormula();
        end = System.currentTimeMillis();

        calcTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells calculate formulas: %.1f s", calcTime.value));

        start = System.currentTimeMillis();
        workbook.save("output/aspose-saved-test-performance.xlsx");
        end = System.currentTimeMillis();
        saveTime.value = (end - start) * 0.001;
        System.out.println(String.format("Aspose.Cells save formulas to Excel: %.1f s", saveTime.value));

        double endMem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
        usedMem.value = (endMem - startMem) / 1024 / 1024;
        System.out.println(String.format("Aspose.Cells used memory: %.1f MB", usedMem.value));

    }
}
