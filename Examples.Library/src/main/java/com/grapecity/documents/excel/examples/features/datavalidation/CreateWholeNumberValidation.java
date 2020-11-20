package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IValidation;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateWholeNumberValidation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {1, 3, 5},
                {7, 9, 11},
                {13, 15, 17}
        });

        //add whole number validation.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.Whole, ValidationAlertStyle.Stop, ValidationOperator.Between, 1, 8);
        IValidation validation = worksheet.getRange("C2:E4").getValidation();
        validation.setIgnoreBlank(true);
        validation.setInputTitle("Tips");
        validation.setInputMessage("Input a value between 1 and 8, please");
        validation.setErrorTitle("Error");
        validation.setErrorMessage("input value does not between 1 and 8");
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
