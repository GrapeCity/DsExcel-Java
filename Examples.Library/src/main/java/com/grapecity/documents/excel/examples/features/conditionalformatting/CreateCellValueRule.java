package com.grapecity.documents.excel.examples.features.conditionalformatting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FormatConditionOperator;
import com.grapecity.documents.excel.FormatConditionType;
import com.grapecity.documents.excel.IFormatCondition;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateCellValueRule extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B:C").setColumnWidth(17);
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

        //weight between 66 and 70, set its interior color to light green.
        IFormatCondition condition = (IFormatCondition) worksheet.getRange("E2:E7").getFormatConditions().add(FormatConditionType.CellValue, FormatConditionOperator.Between, 66, 70);
        condition.getInterior().setColor(Color.GetLightGreen());

    }

}
