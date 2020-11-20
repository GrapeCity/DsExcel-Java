package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.LookAt;
import com.grapecity.documents.excel.ReplaceOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ReplaceWithOptions extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data

        // Skew matrix generator
        // Input:
        // DegX    135
        // DegY    45
        //
        // Output:
        // M11 1	    M12	1
        // M21 -1	M22	1
        // M31 0	    M32	0
        worksheet.getRange("B1").setValue("Skew matrix generator");
        worksheet.getRange("A2:A4").setValue(new String[] {"Input:", "DegX", "DegY"});
        worksheet.getRange("B3").setValue(135);
        worksheet.getRange("B4").setValue(45);
        worksheet.getRange("A6").setValue("Output:");
        worksheet.getRange("A7:A9").setValue(new String[] {"M11", "M21", "M31"});
        worksheet.getRange("B7").setValue(1);
        worksheet.getRange("B8").setFormula("=TAN(B3/180*3.14)");
        worksheet.getRange("B9").setValue(0);
        worksheet.getRange("C7:C9").setValue(new String[] {"M12", "M22", "M32"});
        worksheet.getRange("D7").setFormula("=TAN(B4/180*3.14)");
        worksheet.getRange("D8").setValue(1);
        worksheet.getRange("D9").setValue(0);

        // Replace 3.14 with PI()
        IRange searchRange = worksheet.getUsedRange();
        searchRange.replace(3.14, "PI()");
        
        // Replace M with m (Match case)
        ReplaceOptions tempVar = new ReplaceOptions();
        tempVar.setMatchCase(true);
        searchRange.replace("M", "m", tempVar);
        
        // Replace m11 with M11 (Match whole word, match byte)
        ReplaceOptions tempVar2 = new ReplaceOptions();
        tempVar2.setLookAt(LookAt.Whole);
        tempVar2.setMatchByte(true);
        searchRange.replace("m11", "M11", tempVar2);

	}
}