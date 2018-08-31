package com.grapecity.documents.excel.examples.features.filtering.filteringbyicon;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.AutoFilterOperator;
import com.grapecity.documents.excel.IIconSetCondition;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.IconSetType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FilterByIcon extends ExampleBase {

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

        IIconSetCondition iconset = worksheet.getRange("E2:E7").getFormatConditions().addIconSetCondition();
        iconset.setIconSet(workbook.getIconSets().get(IconSetType.Icon3TrafficLights1));

        //filter in the rows which icon is green light.
        worksheet.getRange("A1:F7").autoFilter(4, workbook.getIconSets().get(IconSetType.Icon3TrafficLights1).get(2), AutoFilterOperator.Icon);

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
