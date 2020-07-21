package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AutoFit extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Auto fit column width of range 'A1'
        worksheet.getRange("A1").setValue("Grapecity Documents for Excel");
        worksheet.getRange("A1").getColumns().autoFit();

        //Auto fit row height of range 'B2'
        worksheet.getRange("B2").setValue("Grapecity");
        worksheet.getRange("B2").getFont().setSize(20);
        worksheet.getRange("B2").getRows().autoFit();

        //Auto fit column width and row height of range 'C3'
        worksheet.getRange("C3").setValue("Grapecity Documents for Excel");
        worksheet.getRange("C3").getFont().setSize(32);
        worksheet.getRange("C3").autoFit();
    }

}
