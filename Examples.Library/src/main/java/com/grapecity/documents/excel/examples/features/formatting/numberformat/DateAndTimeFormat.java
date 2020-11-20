package com.grapecity.documents.excel.examples.features.formatting.numberformat;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DateAndTimeFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:F").setColumnWidth(17);
        worksheet.getRange("A1:F1").setFormula("= Now()");

        // Apply different date formats.
        worksheet.getRange("A1").setNumberFormat("m/d/yy");
        worksheet.getRange("B1").setNumberFormat("d-mmm-yy");
        worksheet.getRange("C1").setNumberFormat("dddd");

        // Apply different time formats.
        worksheet.getRange("D1").setNumberFormat("m/d/yy h:mm");
        worksheet.getRange("E1").setNumberFormat("h:mm AM/PM");
        worksheet.getRange("F1").setNumberFormat("h:mm:ss");
    }

}
