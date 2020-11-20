package com.grapecity.documents.excel.examples.features.conditionalformatting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.DupeUnique;
import com.grapecity.documents.excel.IUniqueValues;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateUniqueRule extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        Object data = new Object[][]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 80, 165},
                {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 72, 134},
                {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 71, 180},
                {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 80, 163},
                {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 58, 176},
                {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 71, 145}
        };
        worksheet.getRange("A1:F7").setValue(data);

        //Unique rule.
        IUniqueValues condition = worksheet.getRange("E2:E7").getFormatConditions().addUniqueValues();
        condition.setDupeUnique(DupeUnique.Unique);
        condition.getFont().setName("Arial");
        condition.getInterior().setColor(Color.GetPink());

    }

}
