package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IValidation;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateListValidation extends ExampleBase {

    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1").setValue("aaa");
        worksheet.getRange("A2").setValue("bbb");
        worksheet.getRange("A3").setValue("ccc");

        worksheet.getRange("C2:E4").setValue(new Object[][]{
                {"aaa", "bbb", "ccc"},
                {"aaa1", "bbb1", "ccc1"},
                {"aaa2", "bbb2", "ccc2"}
        });

        //create list validation.
        worksheet.getRange("C2:E4").getValidation().add(ValidationType.List, ValidationAlertStyle.Stop, ValidationOperator.Between, "=$a$1:$a$3", null);
        IValidation validation = worksheet.getRange("C2:E4").getValidation();
        validation.setInCellDropdown(true);

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
