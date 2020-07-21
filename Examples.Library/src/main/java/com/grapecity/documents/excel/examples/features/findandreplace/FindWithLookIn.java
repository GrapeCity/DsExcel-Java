package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FindLookIn;
import com.grapecity.documents.excel.FindOptions;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SearchDirection;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FindWithLookIn extends ExampleBase {

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
        worksheet.getRange("A3").addComment("Enter the day offset");
        worksheet.getRange("B3").setFormula("=DATE(2019,5,1)");
        worksheet.getRange("B3").setNumberFormat("yyyy-mm-dd;@");
        worksheet.getRange("C3").setFormula("=B3+1");
        worksheet.getRange("C3").setNumberFormat("yyyy-mm-dd;@");
        worksheet.getUsedRange().autoFit();

        // Find the first occurrence of "2019" in the formula bar
        // and mark it with green foreground color

        IRange searchRange = worksheet.getRange("A1:C3");
        FindOptions tempVar = new FindOptions();
        tempVar.setLookIn(FindLookIn.Formulas);

        IRange first2019InFormulaBar = searchRange.find("2019", null, tempVar);
        first2019InFormulaBar.getFont().setColor(Color.GetGreen());

        // Find the first occurrence of 1 in text
        // and mark it with blue foreground
        FindOptions tempVar2 = new FindOptions();
        tempVar2.setLookIn(FindLookIn.Texts);

        IRange firstValue1 = searchRange.find(1, null, tempVar2);
        firstValue1.getFont().setColor(Color.GetBlue());

        // Find the first occurrence of "day" in comments
        // and mark it with yellow background
        FindOptions tempVar3 = new FindOptions();
        tempVar3.setLookIn(FindLookIn.Comments);

        IRange firstDayComments = searchRange.find("day", null, tempVar3);
        firstDayComments.getInterior().setColor(Color.GetYellow());

        // Find the last occurrence of "2019" in the formula property
        // and mark it with purple foreground
        FindOptions tempVar4 = new FindOptions();
        tempVar4.setLookIn(FindLookIn.OnlyFormulas);
        tempVar4.setSearchDirection(SearchDirection.Previous);

        IRange last2019OnlyFormula = searchRange.find("2019", null, tempVar4);
        last2019OnlyFormula.getFont().setColor(Color.GetPurple());
	}
}