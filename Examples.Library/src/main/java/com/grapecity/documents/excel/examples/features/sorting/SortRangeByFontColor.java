package com.grapecity.documents.excel.examples.features.sorting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontColorSortField;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SortOrder;
import com.grapecity.documents.excel.SortOrientation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SortRangeByFontColor extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 67, 165},
                {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 62, 134},
                {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 66, 163},
                {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 68, 145}
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1:F7").setValue(data);
        worksheet.getRange("A:F").setColumnWidth(15);

        worksheet.getRange("F2").getFont().setColor(Color.GetRed());
        worksheet.getRange("F3").getFont().setColor(Color.GetGainsboro());
        worksheet.getRange("F4").getFont().setColor(Color.GetGreen());
        worksheet.getRange("F5").getFont().setColor(Color.GetGainsboro());
        worksheet.getRange("F6").getFont().setColor(Color.GetGainsboro());
        worksheet.getRange("F7").getFont().setColor(Color.GetGainsboro());

        //The font colors which are same to "F7" cell's color will in the top.
        worksheet.getSort().getSortFields().add(new FontColorSortField(worksheet.getRange("F2:F7"), worksheet.getRange("F7").getFont().getColor(), SortOrder.Ascending));
        worksheet.getSort().setRange(worksheet.getRange("A2:F7"));
        worksheet.getSort().setOrientation(SortOrientation.Columns);
        worksheet.getSort().apply();

    }

}
