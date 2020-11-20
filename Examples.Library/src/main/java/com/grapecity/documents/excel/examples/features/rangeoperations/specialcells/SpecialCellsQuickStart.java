package com.grapecity.documents.excel.examples.features.rangeoperations.specialcells;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SpecialCellsQuickStart extends ExampleBase {
    
	@Override
	public void execute(Workbook workbook) {
        IWorksheet ws = workbook.getActiveSheet();

        // Set data
        ws.getRange("A1").setFormula("=\"Text \" & 1");
        ws.getRange("B1").setFormula("=8*10^6");
        ws.getRange("C1").setFormula("=SEARCH(A1,9)");
        ws.getRange("A2").setValue("Text");
        ws.getRange("B2").setValue(1);

        // Find text formulas
        IRange textFormula = ws.getCells().specialCells(SpecialCellType.Formulas, SpecialCellsValue.TextValues);

        // Find number formulas
        IRange numberFormula = ws.getCells().specialCells(SpecialCellType.Formulas, SpecialCellsValue.Numbers);

        // Find error formulas
        IRange errorFormula = ws.getCells().specialCells(SpecialCellType.Formulas, SpecialCellsValue.Errors);

        // Find text values
        IRange textValue = ws.getCells().specialCells(SpecialCellType.Constants, SpecialCellsValue.TextValues);

        // Find number values
        IRange numberValue = ws.getCells().specialCells(SpecialCellType.Constants, SpecialCellsValue.Numbers);

        // Display search result
        ws.getRange("A4:E5").setValue( new Object[][] {
            {"Text formula", "Number Formula", "Error Formula", "Text Value", "Number Value"},
            {textFormula.getAddress(), numberFormula.getAddress(), errorFormula.getAddress(), textValue.getAddress(), numberValue.getAddress()}
        });

		ws.getUsedRange().getEntireColumn().autoFit();
	}

    @Override
    public boolean getIsNew() {
        return true;
    }
}