package com.grapecity.documents.excel.examples.features.slicer;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.ConsolidationFunction;
import com.grapecity.documents.excel.IPivotCache;
import com.grapecity.documents.excel.IPivotField;
import com.grapecity.documents.excel.IPivotTable;
import com.grapecity.documents.excel.ISlicer;
import com.grapecity.documents.excel.ISlicerCache;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PivotFieldOrientation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddSlicersForPivotTable extends ExampleBase {

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
        worksheet.getRange("A1:F16").setValue(sourceData);
        worksheet.getRange("A:F").setColumnWidth(15);

        //Create pivot cache.
        IPivotCache pivotcache = workbook.getPivotCaches().create(worksheet.getRange("A1:F16"));
        //Create pivot tables.
        IPivotTable pivottable1 = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("K5"), "pivottable1");
        IPivotTable pivottable2 = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("N3"), "pivottable2");
        worksheet.getRange("D2:D16").setNumberFormat("$#,##0.00");

        //Config pivot fields
        IPivotField field_product1 = pivottable1.getPivotFields().get(1);
        field_product1.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Amount1 = pivottable1.getPivotFields().get(3);
        field_Amount1.setOrientation(PivotFieldOrientation.DataField);

        IPivotField field_product2 = pivottable2.getPivotFields().get(5);
        field_product2.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Amount2 = pivottable2.getPivotFields().get(2);
        field_Amount2.setOrientation(PivotFieldOrientation.DataField);
        field_Amount2.setFunction(ConsolidationFunction.Count);

        //create slicer cache, the slicers base the slicer cache just control pivot table1.
        ISlicerCache cache = workbook.getSlicerCaches().add(pivottable1, "Product");
        ISlicer slicer1 = cache.getSlicers().add(workbook.getWorksheets().get("Sheet1"), "p1", "Product", 30, 550, 100, 200);

        //add pivot table2 for slicer cache, the slicers base the slicer cache will control pivot tabl1 and pivot table2.
        cache.getPivotTables().addPivotTable(pivottable2);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

}
