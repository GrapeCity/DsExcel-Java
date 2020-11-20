package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateCustomValidation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A2").setValue(1);
        worksheet.getRange("A3").setValue(2);

        worksheet.getRange("C2").setValue(0);

        //create custom validation, if the expression "=$C$2" result is true, the cell's validation will be true, otherwise, it is false.
        //when use custom validation, validationOperator and formula2 parameters will be ignored even if you have given.
        worksheet.getRange("A2:A3").getValidation().add(ValidationType.Custom, ValidationAlertStyle.Information, ValidationOperator.Between, "=$C$2", null);

        //judge if Range["A2:A3"] has validation.
        for (int i = 1; i <= 2; i++) {
            if (worksheet.getRange(i, 0).getHasValidation()) {
                //set the range[i, 0]'s interior color.
                worksheet.getRange(i, 0).getInterior().setColor(Color.GetLightBlue());
            }
        }

    }

}
