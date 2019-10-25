package com.grapecity.documents.excel.examples.features.formulas.precedentsanddependents;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Precedents extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue(100);
        worksheet.getRange("C1").setFormula("=$A$1");
        worksheet.getRange("E1:E5").setFormula("=$A$1");
        
        worksheet.getRange("E2").setFormula("=sum(A1:A2, B4,C1:C3)");
        worksheet.getRange("A1").setValue(1);
        worksheet.getRange("A2").setValue(2);
        worksheet.getRange("B4").setValue(3);
        worksheet.getRange("C1").setValue(4);
        worksheet.getRange("C2").setValue(5);
        worksheet.getRange("C3").setValue(6);
        
        for (IRange item : worksheet.getRange("E2").getPrecedents())
        {
            item.getInterior().setColor(Color.GetPink());
        }
    }

	@Override
	public boolean getIsNew() {
        return true;
    }
}
