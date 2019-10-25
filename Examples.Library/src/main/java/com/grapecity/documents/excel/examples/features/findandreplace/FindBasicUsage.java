package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FindBasicUsage extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data
        final String CorrectWord = "Macro";
        worksheet.getRange("A1:D5").setValue(CorrectWord);

        final String MisspelledWord = "marco";
        worksheet.getRange("A2,C3,D1").setValue(MisspelledWord);

        // Find the first misspelled word
        IRange searchRange = worksheet.getRange("A1:D5");
        IRange firstMisspelled = searchRange.find(MisspelledWord);

        // Mark it with red foreground
        firstMisspelled.getFont().setColor(Color.GetRed());
	}
}