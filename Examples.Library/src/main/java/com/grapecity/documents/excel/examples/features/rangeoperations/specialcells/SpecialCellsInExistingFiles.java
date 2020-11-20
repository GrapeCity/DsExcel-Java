package com.grapecity.documents.excel.examples.features.rangeoperations.specialcells;

import java.io.InputStream;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SpecialCellsInExistingFiles extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        InputStream existingFile = getResourceStream("xlsx/FinancialReport.xlsx");
        workbook.open(existingFile);

        IRange cells = workbook.getActiveSheet().getCells();

        // Find all formulas
        IRange allFormulas = cells.specialCells(SpecialCellType.Formulas);
        // Find all constants
        IRange allConstants = cells.specialCells(SpecialCellType.Constants);

        // Change background color of found cells
        allFormulas.getInterior().setColor(Color.GetLightGray());
        allConstants.getInterior().setColor(Color.GetDarkGray());
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/FinancialReport.xlsx" };
    }
}