package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IValidation;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateDecimalValidation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {1, 3.0, 3.4},
                {7, 9, 102.7},
                {102.8, 110, 120}
        });

        //add decimal validation.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.Decimal, ValidationAlertStyle.Stop, ValidationOperator.Between, 3.4, 102.8);
        IValidation validation = worksheet.getRange("C2:E4").getValidation();
        validation.setIgnoreBlank(true);
        validation.setInputTitle("Tips");
        validation.setInputMessage("Input a decimal value between 3.4 and 102.8, please.");
        validation.setErrorTitle("Error");
        validation.setErrorMessage("input value does not between 3.4 and 102.8.");
        validation.setShowInputMessage(true);
        validation.setShowError(true);

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
