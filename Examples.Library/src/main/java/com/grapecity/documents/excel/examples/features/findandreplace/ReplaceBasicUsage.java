package com.grapecity.documents.excel.examples.features.findandreplace;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ReplaceBasicUsage extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Prepare data
        worksheet.getRange("A1:A3")
                .setValue(new String[] { "Render Excel ranges inside PDF in .NET Core",
                        "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                        "How to format Pivot table styles in .NET Core (Support Team)" });

        // Replace ".NET Core" with ".NET 5"
        worksheet.getUsedRange().replace(".NET Core", ".NET 5");
	}
}