package com.grapecity.documents.excel.examples.features.conditionalformatting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ConditionValueTypes;
import com.grapecity.documents.excel.DataBarAxisPosition;
import com.grapecity.documents.excel.DataBarDirection;
import com.grapecity.documents.excel.DataBarFillType;
import com.grapecity.documents.excel.DataBarNegativeColorType;
import com.grapecity.documents.excel.IDataBar;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateDatabBarRule extends ExampleBase {

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

        //data bar rule.
        IDataBar dataBar = worksheet.getRange("E2:E7").getFormatConditions().addDatabar();

        dataBar.getMinPoint().setType(ConditionValueTypes.LowestValue);
        dataBar.getMinPoint().setValue(null);
        dataBar.getMaxPoint().setType(ConditionValueTypes.HighestValue);
        dataBar.getMaxPoint().setValue(null);

        dataBar.setBarFillType(DataBarFillType.Gradient);
        dataBar.getBarColor().setColor(Color.GetGreen());
        dataBar.setDirection(DataBarDirection.Context);
        dataBar.getAxisColor().setColor(Color.GetRed());
        dataBar.setAxisPosition(DataBarAxisPosition.Automatic);
        dataBar.getNegativeBarFormat().setBorderColorType(DataBarNegativeColorType.Color);
        dataBar.getNegativeBarFormat().getBorderColor().setColor(Color.GetBlue());
        dataBar.getNegativeBarFormat().setColorType(DataBarNegativeColorType.Color);
        dataBar.getNegativeBarFormat().getColor().setColor(Color.GetPink());
        dataBar.setShowValue(false);
    }

}
