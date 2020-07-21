package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.FindOptions;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ReplaceCustomWrapSearch extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data
        worksheet.getRange("A1:A8")
                .setValue(new String[] { "Whats new in GcExcel v2 sp2", "Render Excel ranges inside PDF in .NET Core",
                        "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                        "How to format Pivot table styles in .NET Core (Support Team)",
                        "Controlling page breaks when editing Excel files in .NET Core (Support Team)",
                        "Combine different workbooks into PDF in .NET Core (Support Team)",
                        "Repeating Excel rows/columns on exporting to PDF in .NET Core (Support Team)",
                        "Using GcExcel with Kotlin" });

        // Find ".NET Core" and replace them with ".NET 5", starting after A4
        String what = ".NET Core";
        String replacement = ".NET 5";
        FindOptions settings = new FindOptions();
        IRange target = worksheet.getUsedRange();
        IRange after = worksheet.getRange("A4");

        // Search start after A4
        IRange cellToReplace = after;
        do {
            cellToReplace = target.find(what, cellToReplace, settings);
            if (cellToReplace == null) {
                break;
            }

            // Replace
            cellToReplace.setValue(cellToReplace.getText().replace(what, replacement));
        } while (true);

        // Search reached the bottom of the range.
        // Wrap search start at the top-left corner.
        if (after != null) {
            do {
                cellToReplace = target.find(what, cellToReplace, settings);
                if (cellToReplace == null) {
                    break;
                }

                // Replace
                cellToReplace.setValue(cellToReplace.getText().replace(what, replacement));

                if (cellToReplace.getRow() == after.getRow() && cellToReplace.getColumn() == after.getColumn()) {
                    break;
                }
            } while (true);
        }
	}
}