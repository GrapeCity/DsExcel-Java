package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.util.GregorianCalendar;

public class CreateTimeValidation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {new GregorianCalendar(1899, 11, 30, 13, 0, 0), new GregorianCalendar(1899, 11, 30, 13, 29, 59), new GregorianCalendar(1899, 11, 30, 13, 30, 0)},
                {new GregorianCalendar(1899, 11, 30, 14, 0, 0), new GregorianCalendar(1899, 11, 30, 15, 0, 0), new GregorianCalendar(1899, 11, 30, 16, 30, 0)},
                {new GregorianCalendar(1899, 11, 30, 19, 0, 0), new GregorianCalendar(1899, 11, 30, 18, 29, 59), new GregorianCalendar(1899, 11, 30, 18, 30, 0)}
        });

        //create time validation.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.Time, ValidationAlertStyle.Stop, ValidationOperator.Between, new GregorianCalendar(1899, 11, 30, 13, 30, 0), new GregorianCalendar(1899, 11, 30, 18, 30, 0));

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
