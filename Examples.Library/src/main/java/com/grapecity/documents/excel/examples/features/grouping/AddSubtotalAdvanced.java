package com.grapecity.documents.excel.examples.features.grouping;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddSubtotalAdvanced extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IRange targetRange = workbook.getActiveSheet().getRange("A1:C9");
        // Set data
        targetRange.setValue(new Object[][] { { "Grade", "Class", "Score", "Student ID" }, { 1, 1, 93, 1 },
                { 1, 1, 87, 2 }, { 1, 2, 97, 3 }, { 1, 2, 95, 4 }, { 2, 1, 83, 5 }, { 2, 1, 87, 6 }, { 2, 2, 96, 7 },
                { 2, 2, 83, 8 } });

        // Group by Grade select Average(Score)
        targetRange.subtotal(1, // Grade
                ConsolidationFunction.Average, new int[] { 3 }, // Score
                false, true);

        // Group by Class select Average(Score)
        targetRange.subtotal(2, // Class
                ConsolidationFunction.Average, new int[] { 3 }, // Score
                false);

        workbook.getActiveSheet().getRange("C:C").setNumberFormat("0;0;0;@");

        targetRange.autoFit();
    }
}
