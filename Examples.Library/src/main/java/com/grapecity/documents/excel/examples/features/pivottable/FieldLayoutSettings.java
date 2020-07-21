package com.grapecity.documents.excel.examples.features.pivottable;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.util.GregorianCalendar;

public class FieldLayoutSettings extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object sourceData = new Object[][]{
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Bose 785593-0050", "Consumer Electronics", 4270, new GregorianCalendar(2018, 0, 6), "United States"},
                {2, "Canon EOS 1500D", "Consumer Electronics", 8239, new GregorianCalendar(2018, 0, 7), "United Kingdom"},
                {3, "Haier 394L 4Star", "Consumer Electronics", 617, new GregorianCalendar(2018, 0, 8), "United States"},
                {4, "IFB 6.5 Kg FullyAuto", "Consumer Electronics", 8384, new GregorianCalendar(2018, 0, 10), "Canada"},
                {5, "Mi LED 40inch", "Consumer Electronics", 2626, new GregorianCalendar(2018, 0, 10), "Germany"},
                {6, "Sennheiser HD 4.40-BT", "Consumer Electronics", 3610, new GregorianCalendar(2018, 0, 11), "United States"},
                {7, "Iphone XR", "Mobile", 9062, new GregorianCalendar(2018, 0, 11), "Australia"},
                {8, "OnePlus 7Pro", "Mobile", 6906, new GregorianCalendar(2018, 0, 16), "New Zealand"},
                {9, "Redmi 7", "Mobile", 2417, new GregorianCalendar(2018, 0, 16), "France"},
                {10, "Samsung S9", "Mobile", 7431, new GregorianCalendar(2018, 0, 16), "Canada"},
                {11, "OnePlus 7Pro", "Mobile", 8250, new GregorianCalendar(2018, 0, 16), "Germany"},
                {12, "Redmi 7", "Mobile", 7012, new GregorianCalendar(2018, 0, 18), "United States"},
                {13, "Bose 785593-0050", "Consumer Electronics", 1903, new GregorianCalendar(2018, 0, 20), "Germany"},
                {14, "Canon EOS 1500D", "Consumer Electronics", 2824, new GregorianCalendar(2018, 0, 22), "Canada"},
                {15, "Haier 394L 4Star", "Consumer Electronics", 6946, new GregorianCalendar(2018, 0, 24), "France"},
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("G1:L16").setValue(sourceData);
        worksheet.getRange("G:L").setColumnWidth(15);
        IPivotCache pivotcache = workbook.getPivotCaches().create(worksheet.getRange("G1:L16"));
        IPivotTable pivottable = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("A1"), "pivottable1");
        worksheet.getRange("J1:J16").setNumberFormat("$#,##0.00");

        //config pivot table's fields
        IPivotField field_Category = pivottable.getPivotFields().get("Category");
        field_Category.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Product = pivottable.getPivotFields().get("Product");
        field_Product.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Country = pivottable.getPivotFields().get("Country");
        field_Country.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Amount = pivottable.getPivotFields().get("Amount");
        field_Amount.setOrientation(PivotFieldOrientation.DataField);
        field_Amount.setNumberFormat("$#,##0.00");

        // tablular row
        field_Category.setLayoutForm(LayoutFormType.Tabular);
        // insert blank row
        field_Category.setLayoutBlankLine(true);

        // outline row
        field_Country.setLayoutForm(LayoutFormType.Outline);
        field_Country.setLayoutCompactRow(false);
        // show subtotallocation on bottom
        field_Country.setLayoutSubtotalLocation(SubtotalLocationType.Bottom);

        worksheet.getRange("A:C").getEntireColumn().autoFit();
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