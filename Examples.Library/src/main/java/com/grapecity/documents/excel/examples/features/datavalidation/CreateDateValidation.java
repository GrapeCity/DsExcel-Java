package com.grapecity.documents.excel.examples.features.datavalidation;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateDateValidation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {new GregorianCalendar(2020, 11, 1), new GregorianCalendar(2020, 11, 14), new GregorianCalendar(2020, 11, 15)},
                {new GregorianCalendar(2020, 11, 18), new GregorianCalendar(2020, 11, 19), new GregorianCalendar(2020, 11, 30)},
                {new GregorianCalendar(2020, 11, 31), new GregorianCalendar(2019, 11, 13), new GregorianCalendar(2019, 11, 15)},
        });

        //create date validation.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.Date, ValidationAlertStyle.Stop, ValidationOperator.Between, new GregorianCalendar(2020, 11, 13), new GregorianCalendar(2020, 11, 18));

        //set column width just for export shown.
        worksheet.getRange("C:E").getEntireColumn().setColumnWidthInPixel(120);

        //judge if Range["C2:E4"] has validation.
        for (int i = 1; i <= 3; i++) {
            for (int j = 2; j <= 4; j++) {
                if (worksheet.getRange(i, j).getHasValidation()) {
                    //set the range[i, j]'s interior color.
                    worksheet.getRange(i, j).getInterior().setColor(Color.GetLightBlue());
                }
            }
        }

    }

}
