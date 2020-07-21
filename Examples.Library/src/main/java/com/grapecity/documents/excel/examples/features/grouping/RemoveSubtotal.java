package com.grapecity.documents.excel.examples.features.grouping;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class RemoveSubtotal extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IRange targetRange = workbook.getActiveSheet().getRange("A1:C9");
        // Set data
        targetRange.setValue(new Object[][] { { "Player", "Side", "Commander" }, { 1, "Soviet", "AI" },
                { 2, "Soviet", "AI" }, { 3, "Soviet", "Human" }, { 4, "Allied", "Human" }, { 5, "Allied", "Human" },
                { 6, "Allied", "AI" }, { 7, "Empire", "AI" }, { 8, "Empire", "AI" } });

        // Subtotal
        targetRange.subtotal(2, // Side
                ConsolidationFunction.Count, new int[] { 2 } // Side
        );

        workbook.getActiveSheet().getRange("A1:C13").removeSubtotal();
    }
}
