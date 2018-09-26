package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateTextLength extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {"aa", "bb", "aa1"},
                {"ccc", "dddd", "dddd1"},
                {"ff", "ffff", "ffff1"}
        });

        //create text length validation, text length between 2 and 3.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.TextLength, ValidationAlertStyle.Stop, ValidationOperator.Between, 2, 3);
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
