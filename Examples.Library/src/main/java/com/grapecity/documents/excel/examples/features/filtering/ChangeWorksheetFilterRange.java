package com.grapecity.documents.excel.examples.features.filtering;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ChangeWorksheetFilterRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("D3").setValue("Numbers");
        worksheet.getRange("D4").setValue(1);
        worksheet.getRange("D5").setValue(2);
        worksheet.getRange("D6").setValue(3);

        //first time invoke. worksheet's filter range will be D3:D6.
        worksheet.getRange("D3:D6").autoFilter(0, "<>2");

        //set AutoFilterMode to false. must set this property before, otherwise, subsequent code can not take effort.
        worksheet.setAutoFilterMode(false);

        worksheet.getRange("A5").setValue("Numbers");
        worksheet.getRange("A6").setValue(1);
        worksheet.getRange("A7").setValue(2);
        worksheet.getRange("A8").setValue(3);

        //second time invoke. worksheet's filter range will change to A5:A8.
        worksheet.getRange("A5:A8").autoFilter(0, "<>2");

    }

}
