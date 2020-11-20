package com.grapecity.documents.excel.examples.features.workbook.options;

import com.grapecity.documents.excel.IPivotCache;
import com.grapecity.documents.excel.IPivotTable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PivotFieldOrientation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.util.GregorianCalendar;

public class ConfigAutomaticGroupDateTimeInPivotTable extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        Object sourceData = new Object[][]{
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Bose 785593-0050", "Consumer Electronics", 4270, new GregorianCalendar(2011, 0, 6), "United States"},
                {2, "Canon EOS 1500D", "Consumer Electronics", 8239, new GregorianCalendar(2011, 0, 7), "United Kingdom"},
                {3, "Haier 394L 4Star", "Consumer Electronics", 617, new GregorianCalendar(2011, 5, 8), "United States"},
                {4, "IFB 6.5 Kg FullyAuto", "Consumer Electronics", 8384, new GregorianCalendar(2012, 0, 10), "Canada"},
                {5, "Mi LED 40inch", "Consumer Electronics", 2626, new GregorianCalendar(2012, 1, 10), "Germany"},
                {6, "Sennheiser HD 4.40-BT", "Consumer Electronics", 3610, new GregorianCalendar(2012, 0, 11), "United States"},
                {7, "Iphone XR", "Mobile", 9062, new GregorianCalendar(2011, 0, 11), "Australia"},
                {8, "OnePlus 7Pro", "Mobile", 6906, new GregorianCalendar(2012, 4, 16), "New Zealand"},
                {9, "Redmi 7", "Mobile", 2417, new GregorianCalendar(2012, 0, 16), "France"},
                {10, "Samsung S9", "Mobile", 7431, new GregorianCalendar(2012, 0, 16), "Canada"},
                {11, "OnePlus 7Pro", "Mobile", 8250, new GregorianCalendar(2013, 0, 16), "Germany"},
                {12, "Redmi 7", "Mobile", 7012, new GregorianCalendar(2012, 2, 18), "United States"},
                {13, "Bose 785593-0050", "Consumer Electronics", 1903, new GregorianCalendar(2013, 0, 20), "Germany"},
                {14, "Canon EOS 1500D", "Consumer Electronics", 2824, new GregorianCalendar(2012, 3, 22), "Canada"},
                {15, "Haier 394L 4Star", "Consumer Electronics", 6946, new GregorianCalendar(2013, 0, 24), "France"},
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set not to group date/time fileds in PivotTable automatically.
        workbook.getOptions().getData().setAutomaticGroupDateTimeInPivotTable(false);

        worksheet.getRange("Q1:V16").setValue(sourceData);
        IPivotCache pivotcache = workbook.getPivotCaches().create(worksheet.getRange("Q1:V16"));
        IPivotTable pivottable = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("A1"), "pivottable1");

        pivottable.getPivotFields().get("Category").setOrientation(PivotFieldOrientation.RowField);
        pivottable.getPivotFields().get("Product").setOrientation(PivotFieldOrientation.RowField);
        pivottable.getPivotFields().get("Date").setOrientation(PivotFieldOrientation.ColumnField);
        pivottable.getPivotFields().get("Amount").setOrientation(PivotFieldOrientation.DataField);

        worksheet.getRange("A:P").getEntireColumn().autoFit();
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getShowScreenshot() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
