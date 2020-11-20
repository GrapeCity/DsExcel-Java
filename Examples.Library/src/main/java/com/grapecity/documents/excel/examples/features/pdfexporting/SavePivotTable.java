package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.IPivotCache;
import com.grapecity.documents.excel.IPivotField;
import com.grapecity.documents.excel.IPivotTable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PivotFieldOrientation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SavePivotTable extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        
        Object sourceData = new Object[][]{
            {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
            {1, "Broccoli", "Vegetables", 8239, new GregorianCalendar(2018, 0, 7), "United Kingdom"},
            {2, "Banana", "Fruit", 617, new GregorianCalendar(2018, 0, 8), "United States"},
            {3, "Banana", "Fruit", 8384, new GregorianCalendar(2018, 0, 10), "Canada"},
            {4, "Beans", "Vegetables", 2626, new GregorianCalendar(2018, 0, 10), "Germany"},
            {5, "Orange", "Fruit", 3610, new GregorianCalendar(2018, 0, 11), "United States"},
            {6, "Broccoli", "Vegetables", 9062, new GregorianCalendar(2018, 0, 11), "Australia"},
            {7, "Banana", "Fruit", 6906, new GregorianCalendar(2018, 0, 16), "New Zealand"},
            {8, "Apple", "Fruit", 2417, new GregorianCalendar(2018, 0, 16), "France"},
            {9, "Apple", "Fruit", 7431, new GregorianCalendar(2018, 0, 16), "Canada"},
            {10, "Banana", "Fruit", 8250, new GregorianCalendar(2018, 0, 16), "Germany"},
            {11, "Broccoli", "Vegetables", 7012, new GregorianCalendar(2018, 0, 18), "United States"},
            {12, "Broccoli", "Vegetables", 2824, new GregorianCalendar(2018, 0, 22), "Canada"},
            {13, "Apple", "Fruit", 6946, new GregorianCalendar(2018, 0, 24), "France"},
         };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("K20:P33").setValue(sourceData);
        worksheet.getRange("K:P").setColumnWidth(15);
        // Add pivot table
        IPivotCache pivotcache = workbook.getPivotCaches().create(worksheet.getRange("K20:P33"));
        IPivotTable pivottable = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("A1"), "pivottable1");
        worksheet.getRange("N21:N35").setNumberFormat("$#,##0.00");
        worksheet.getRange("A:G").setColumnWidth(12);
        
        // config pivot table's fields
        IPivotField field_Date = pivottable.getPivotFields().get("Date");
        field_Date.setOrientation(PivotFieldOrientation.PageField);

        IPivotField field_Category = pivottable.getPivotFields().get("Category");
        field_Category.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Product = pivottable.getPivotFields().get("Product");
        field_Product.setOrientation(PivotFieldOrientation.ColumnField);

        IPivotField field_Amount = pivottable.getPivotFields().get("Amount");
        field_Amount.setOrientation(PivotFieldOrientation.DataField);
        field_Amount.setNumberFormat("$#,##0.00");

        IPivotField field_Country = pivottable.getPivotFields().get("Country");
        field_Country.setOrientation(PivotFieldOrientation.RowField);

        // Set pivot style
        pivottable.setTableStyle("PivotStyleMedium28");
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

}
