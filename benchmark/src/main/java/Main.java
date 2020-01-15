import com.grapecity.documents.excel.*;
import java.io.*;

public class Main {
    public static void main(String[] args) throws Exception {

        // the reason I only give 100,000*30 cells is that POI consumes nearly 4G memory in current benchmark, it will be out of memory
        int rowCount = 100000;
        int colCount = 30;

        String product = "";
        String scenario = "double";

        if (args.length != 0) {
            //product = args[0];
            scenario = args[0];
        }

        RefObject<Double> setTime = new RefObject<>();
        RefObject<Double> getTime = new RefObject<>();
        RefObject<Double> saveTime = new RefObject<>();
        RefObject<Double> usedMem = new RefObject<>();

        Workbook resultBook = new Workbook();
        File benchmarks = new File("output/benchmarks.xlsx");
        if(benchmarks.exists()){
            resultBook.open("output/benchmarks.xlsx");
        }

        String scenarioLower = scenario.toLowerCase();
        if (scenarioLower.equals("double")) {

            IWorksheet doublesSheet = createResutlSheet(resultBook,"double");

            GcExcelBenchmark.TestSetRangeValues_Double(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "gcexcel", setTime, getTime, saveTime, usedMem);

            AsposeBenchmark.TestSetRangeValues_Double(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "aspose", setTime, getTime, saveTime, usedMem);

            POIBenchmark.TestSetRangeValues_Double(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "poi", setTime, getTime, saveTime, usedMem);

        } else if (scenarioLower.equals("string")) {
            IWorksheet doublesSheet = createResutlSheet(resultBook,"string");

            GcExcelBenchmark.TestSetRangeValues_String(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "gcexcel", setTime, getTime, saveTime, usedMem);

            AsposeBenchmark.TestSetRangeValues_String(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "aspose", setTime, getTime, saveTime, usedMem);

            POIBenchmark.TestSetRangeValues_String(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "poi", setTime, getTime, saveTime, usedMem);

        } else if (scenarioLower.equals("date")) {
            IWorksheet doublesSheet = createResutlSheet(resultBook,"date");

            GcExcelBenchmark.TestSetRangeValues_Date(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "gcexcel", setTime, getTime, saveTime, usedMem);

            AsposeBenchmark.TestSetRangeValues_Date(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "aspose", setTime, getTime, saveTime, usedMem);

            POIBenchmark.TestSetRangeValues_Date(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "poi", setTime, getTime, saveTime, usedMem);

        } else if (scenarioLower.equals("bigfile")) {
            IWorksheet doublesSheet = createResutlSheet2(resultBook,"big file");

            RefObject<Double> openTime = new RefObject<>();
            RefObject<Double> calcTime = new RefObject<>();

            GcExcelBenchmark.TestBigExcelFile(rowCount, colCount, openTime, calcTime, saveTime, usedMem);
            FillResult(doublesSheet, "gcexcel", openTime, calcTime, saveTime, usedMem);

            AsposeBenchmark.TestBigExcelFile(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "aspose", openTime, calcTime, saveTime, usedMem);

            POIBenchmark.TestBigExcelFile(rowCount, colCount, openTime, calcTime, saveTime, usedMem);
            FillResult(doublesSheet, "poi", openTime, calcTime, saveTime, usedMem);
            
        }else if(scenario.equals("formula")){
            IWorksheet doublesSheet = createResutlSheet3(resultBook,"formulas");

            // Aspose can only handle 20000*30 formulas, the time of saving Excel will be very long when the amount of formulas is bigger than 20000*30
            rowCount = 20000;
            GcExcelBenchmark.TestSetRangeFormulas(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "gcexcel", setTime, getTime, saveTime, usedMem);

            POIBenchmark.TestSetRangeFormulas(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "aspose", setTime, getTime, saveTime, usedMem);

            AsposeBenchmark.TestSetRangeFormulas(rowCount, colCount, setTime, getTime, saveTime, usedMem);
            FillResult(doublesSheet, "poi", setTime, getTime, saveTime, usedMem);
        }

        resultBook.save("output/benchmarks.xlsx");
    }

    private static IWorksheet createResutlSheet(IWorkbook workbook, String name){
        IWorksheet resultSheet = workbook.getWorksheets().get(name);
        if(resultSheet == null){
            resultSheet = workbook.getWorksheets().add();
            resultSheet.setName(name);

            resultSheet.getRange("B1:D1").setValue(new Object[][]{
                {"GcExcel", "POI", "Aspose"}
            });
            resultSheet.getRange("A2:A5").setValue(new Object[][]{
                {"SetTime"},
                {"GetTime"},
                {"SaveTime"},
                {"UsedMem"}
            });

            resultSheet.getNames().add("gcexcel", "B2:B5");
            resultSheet.getNames().add("poi", "C2:C5");
            resultSheet.getNames().add("aspose", "D2:D5");
        }


        return resultSheet;
    }

    private static IWorksheet createResutlSheet2(IWorkbook workbook, String name){
        IWorksheet resultSheet = workbook.getWorksheets().get(name);
        if(resultSheet == null){
            resultSheet = workbook.getWorksheets().add();
            resultSheet.setName(name);

            resultSheet.getRange("B1:D1").setValue(new Object[][]{
                {"GcExcel", "POI", "Aspose"}
            });
            resultSheet.getRange("A2:A5").setValue(new Object[][]{
                {"OpenTim"},
                {"CalcTime"},
                {"SaveTime"},
                {"UsedMem"}
            });

            resultSheet.getNames().add("gcexcel", "B2:B5");
            resultSheet.getNames().add("poi", "C2:C5");
            resultSheet.getNames().add("aspose", "D2:D5");
        }


        return resultSheet;
    }

    private static IWorksheet createResutlSheet3(IWorkbook workbook, String name){
        IWorksheet resultSheet = workbook.getWorksheets().get(name);
        if(resultSheet == null){
            resultSheet = workbook.getWorksheets().add();
            resultSheet.setName(name);

            resultSheet.getRange("B1:D1").setValue(new Object[][]{
                {"GcExcel", "POI", "Aspose"}
            });
            resultSheet.getRange("A2:A5").setValue(new Object[][]{
                {"SetTime"},
                {"CalcTime"},
                {"SaveTime"},
                {"UsedMem"}
            });

            resultSheet.getNames().add("gcexcel", "B2:B5");
            resultSheet.getNames().add("poi", "C2:C5");
            resultSheet.getNames().add("aspose", "D2:D5");
        }

        return resultSheet;
    }

    public static void FillResult(IWorksheet worksheet, String product,
                                                 RefObject<Double> setTime,
                                                 RefObject<Double> getTime,
                                                 RefObject<Double> saveTime,
                                                 RefObject<Double> usedMem) {

        String refersTo = worksheet.getNames().get(product).getRefersTo();
        String range = refersTo.substring(refersTo.indexOf("\"") + 1, refersTo.lastIndexOf("\""))  ;
        worksheet.getRange(range).setValue(new double[][]{
            {setTime.value},
            {getTime.value},
            {saveTime.value},
            {usedMem.value}
        });

    }

}
