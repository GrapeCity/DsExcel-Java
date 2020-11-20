package com.grapecity.documents.excel.examples.features.conditionalformatting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ColorScaleType;
import com.grapecity.documents.excel.ConditionValueTypes;
import com.grapecity.documents.excel.IColorScale;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateThreeColorScaleRule extends ExampleBase {

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

        //three color scale.
        IColorScale threeColorScaleRule = worksheet.getRange("E2:E7").getFormatConditions().addColorScale(ColorScaleType.ThreeColorScale);
        threeColorScaleRule.getColorScaleCriteria().get(0).setType(ConditionValueTypes.LowestValue);
        threeColorScaleRule.getColorScaleCriteria().get(0).getFormatColor().setColor(Color.GetRed());

        threeColorScaleRule.getColorScaleCriteria().get(1).setType(ConditionValueTypes.Percent);
        threeColorScaleRule.getColorScaleCriteria().get(1).setValue(30);
        threeColorScaleRule.getColorScaleCriteria().get(1).getFormatColor().setColor(Color.GetYellow());

        threeColorScaleRule.getColorScaleCriteria().get(2).setType(ConditionValueTypes.HighestValue);
        threeColorScaleRule.getColorScaleCriteria().get(2).getFormatColor().setColor(Color.GetGreen());

    }

}
