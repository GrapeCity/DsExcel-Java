package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FindOptions;
import com.grapecity.documents.excel.IDisplayFormat;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FindDisplayFormat extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data
        worksheet.getRange("A1:C3").setValue("Text");

        IRange b2 = worksheet.getRange("B2");
        b2.getInterior().setColor(Color.GetRed());
        b2.getFont().setColor(Color.GetWhite());
        b2.setValue("B2");

        IRange a2 = worksheet.getRange("A2");
        a2.getInterior().setColor(Color.GetOrange());
        a2.getFont().setColor(Color.GetWhite());
        a2.setValue("A2");

        // Find cells with red background and white foreground,
        // and highlight them with bold and bigger text

        // Create a temporary sheet to build a IDisplayFormat
        IWorksheet displayFormatFactoryWorksheet = workbook.getWorksheets().add();
        IRange displayFormatFactoryRange = displayFormatFactoryWorksheet.getRange("A1");
        displayFormatFactoryRange.getInterior().setColor(Color.GetRed());
        displayFormatFactoryRange.getFont().setColor(Color.GetWhite());
        IDisplayFormat searchFormat = displayFormatFactoryRange.getDisplayFormat();

        // Find the first occurrence
        IRange searchRange = worksheet.getUsedRange();
        FindOptions options = new FindOptions();
        options.setSearchFormat(searchFormat);
        IRange foundCell = searchRange.find("*", null, options);

        // Highlight the found range
        foundCell.getFont().setBold(true);
        foundCell.getFont().setSize(foundCell.getFont().getSize() + 8);

        // Dispose the temporary sheet
        displayFormatFactoryWorksheet.delete();
	}
}