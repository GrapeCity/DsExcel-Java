package com.grapecity.documents.excel.examples.features.filtering.datefiltering;

import java.util.Date;
import java.util.GregorianCalendar;

import com.grapecity.documents.excel.AutoFilterOperator;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DateFilter extends ExampleBase {

    public static void main(String[] args){
        Workbook workbook = new Workbook();
        new DateFilter().execute(workbook);
        workbook.save("/Users/jackshang/Documents/Work/Source/GcExcel/ds-excel-java/Source/Demos/ExamplesDemo/Examples.Library/src/main/java/com/grapecity/documents/excel/examples/features/filtering/datefiltering/dateFilter.xlsx");
    }

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        Object data = new Object[][]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 67, 165},
                {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 62, 134},
                {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 66, 163},
                {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 68, 145}
        };

        worksheet.getRange("A1:F7").setValue(data);
        worksheet.getRange("A:F").setColumnWidth(15);

        String criteria1 = new GregorianCalendar(1972, 6, 3).getTime().toString();
        String criteria2 = new GregorianCalendar(1993, 1, 15).getTime().toString();
        //filter date between 1972.7.3 and 1993.2.15
        worksheet.getRange("A1:F7").autoFilter(2, ">=" + criteria1, AutoFilterOperator.And, "<=" + criteria2);

    }

}
