package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FindLookIn;
import com.grapecity.documents.excel.FindOptions;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.LookAt;
import com.grapecity.documents.excel.SearchDirection;
import com.grapecity.documents.excel.SearchOrder;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FindLookAtSearchOrderDirectionMatchCase extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data

        // Add day to date
        // Day Date Result
        // 1 2019-05-01 2019-05-02
        worksheet.getRange("A2:C2").setValue(new String[] { "Day", "Date", "Result" });
        worksheet.getRange("A1").setValue("Add day to date");
        worksheet.getRange("A3").setValue(1);
        worksheet.getRange("B3").setFormula("=DATE(2019,5,1)");
        worksheet.getRange("B3").setNumberFormat("yyyy-mm-dd;@");
        worksheet.getRange("C3").setFormula("=B3+1");
        worksheet.getRange("C3").setNumberFormat("yyyy-mm-dd;@");
        worksheet.getUsedRange().autoFit();

        IRange searchRange = worksheet.getRange("A1:C3");

        // Find the last occurrence of 1 in text (match whole word, backward and by
        // columns)
        // and mark it with blue foreground and bigger font
        FindOptions tempVar = new FindOptions();
        tempVar.setLookIn(FindLookIn.Texts);
        tempVar.setSearchDirection(SearchDirection.Previous);
        tempVar.setLookAt(LookAt.Whole);
        tempVar.setSearchOrder(SearchOrder.ByColumns);

        IRange lastValue1 = searchRange.find(1, null, tempVar);
        lastValue1.getFont().setColor(Color.GetBlue());
        lastValue1.getFont().setSize(lastValue1.getFont().getSize() + 8);

	}
}