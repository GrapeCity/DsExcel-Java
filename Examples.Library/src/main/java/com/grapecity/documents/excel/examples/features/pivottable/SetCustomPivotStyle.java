package com.grapecity.documents.excel.examples.features.pivottable;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IPivotCache;
import com.grapecity.documents.excel.IPivotField;
import com.grapecity.documents.excel.IPivotTable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PivotFieldOrientation;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetCustomPivotStyle extends ExampleBase {

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
        worksheet.getRange("K20:P33").setValue(sourceData);
        worksheet.getRange("K:P").setColumnWidth(15);
        // Add pivot table
        IPivotCache pivotcache = workbook.getPivotCaches().create(worksheet.getRange("K20:P33"));
        IPivotTable pivottable = worksheet.getPivotTables().add(pivotcache, worksheet.getRange("A1"), "pivottable1");
        worksheet.getRange("N21:N35").setNumberFormat("$#,##0.00");

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

        // Create pivot style "test"
        ITableStyle pivotStyle = workbook.getTableStyles().add("test");

        // Set table style as pivot table style
        pivotStyle.setShowAsAvailablePivotStyle(true);

        pivotStyle.getTableStyleElements().get(TableStyleElementType.PageFieldLabels).getInterior().setColor(Color.GetLightBlue());
        pivotStyle.getTableStyleElements().get(TableStyleElementType.PageFieldValues).getInterior().setColor(Color.GetLightBlue());

        pivotStyle.getTableStyleElements().get(TableStyleElementType.GrandTotalColumn).getInterior().setColor(Color.GetLightGreen());
        pivotStyle.getTableStyleElements().get(TableStyleElementType.GrandTotalRow).getInterior().setColor(Color.GetLightGreen());

        pivotStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setColor(Color.GetCyan());
        pivotStyle.getTableStyleElements().get(TableStyleElementType.FirstColumn).getInterior().setColor(Color.GetTomato());

        pivotStyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe).getInterior().setColor(Color.GetYellow());
        pivotStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).getInterior().setColor(Color.GetLightYellow());

        // Set ShowTableStyleRowStripes as true
        pivottable.setShowTableStyleRowStripes(true);

        // Set pivot table style
        pivottable.setStyle(pivotStyle);

        worksheet.getRange("A1:H16").getColumns().autoFit();
	}

}
