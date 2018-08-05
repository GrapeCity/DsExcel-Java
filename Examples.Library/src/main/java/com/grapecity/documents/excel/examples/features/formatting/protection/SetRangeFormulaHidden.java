package com.grapecity.documents.excel.examples.features.formatting.protection;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetRangeFormulaHidden extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B1").setFormula("=A1");

        // config range B1's FormulaHidden property.
        worksheet.getRange("B1").setFormulaHidden(true);
        // protect worksheet, range B1's formula will not show in exported xlsx file.
        worksheet.setProtection(true);
    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
