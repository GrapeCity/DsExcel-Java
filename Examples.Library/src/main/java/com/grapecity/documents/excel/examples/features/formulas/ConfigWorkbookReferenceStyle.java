package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.ReferenceStyle;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigWorkbookReferenceStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        //set workbook's reference style to R1C1. exported xlsx file will be R1C1 style.
        workbook.setReferenceStyle(ReferenceStyle.R1C1);
    }

}
