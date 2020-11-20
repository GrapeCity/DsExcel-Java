package com.grapecity.documents.excel.examples.features.slicer;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.ISlicer;
import com.grapecity.documents.excel.ISlicerCache;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddSlicersForTable extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object sourceData = new Object[][]{
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Carrots", "Vegetables", 4270, new GregorianCalendar(2018, 0, 6), "United States"},
                {2, "Broccoli", "Vegetables", 8239, new GregorianCalendar(2018, 0, 7), "United Kingdom"},
                {3, "Banana", "Fruit", 617, new GregorianCalendar(2018, 0, 8), "United States"},
                {4, "Banana", "Fruit", 8384, new GregorianCalendar(2018, 0, 10), "Canada"},
                {5, "Beans", "Vegetables", 2626, new GregorianCalendar(2018, 0, 10), "Germany"},
                {6, "Orange", "Fruit", 3610, new GregorianCalendar(2018, 0, 11), "United States"},
                {7, "Broccoli", "Vegetables", 9062, new GregorianCalendar(2018, 0, 11), "Australia"},
                {8, "Banana", "Fruit", 6906, new GregorianCalendar(2018, 0, 16), "New Zealand"},
                {9, "Apple", "Fruit", 2417, new GregorianCalendar(2018, 0, 16), "France"},
                {10, "Apple", "Fruit", 7431, new GregorianCalendar(2018, 0, 16), "Canada"},
                {11, "Banana", "Fruit", 8250, new GregorianCalendar(2018, 0, 16), "Germany"},
                {12, "Broccoli", "Vegetables", 7012, new GregorianCalendar(2018, 0, 18), "United States"},
                {13, "Carrots", "Vegetables", 1903, new GregorianCalendar(2018, 0, 20), "Germany"},
                {14, "Broccoli", "Vegetables", 2824, new GregorianCalendar(2018, 0, 22), "Canada"},
                {15, "Apple", "Fruit", 6946, new GregorianCalendar(2018, 0, 24), "France"},
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:F").setColumnWidth(15);

        worksheet.getRange("A1:F16").setValue(sourceData);
        ITable table = worksheet.getTables().add(worksheet.getRange("A1:F16"), true);
        table.getColumns().get(3).getDataBodyRange().setNumberFormat("$#,##0.00");

        //create slicer cache for table.
        ISlicerCache cache = workbook.getSlicerCaches().add(table, "Category", "categoryCache");

        //add two slicers for Category column.
        ISlicer slicer1 = cache.getSlicers().add(workbook.getWorksheets().get("Sheet1"), "cate1", "Category", 30, 550, 100, 200);
        ISlicer slicer2 = cache.getSlicers().add(workbook.getWorksheets().get("Sheet1"), "cate2", "Category", 30, 700, 100, 200);

    }

}
