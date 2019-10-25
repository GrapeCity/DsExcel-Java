package com.grapecity.documents.excel.examples.features.formulas.precedentsanddependents;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Dependents extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue(100);
        worksheet.getRange("C1").setFormula("=$A$1");
        worksheet.getRange("E1:E5").setFormula("=$A$1");
        for (IRange item : worksheet.getRange("A1").getDependents())
        {
            item.getInterior().setColor(Color.GetAzure());
        }
    }

	@Override
	public boolean getIsNew() {
        return true;
    }
}
