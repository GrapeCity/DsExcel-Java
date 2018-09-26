package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateNewWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Create empty workbook, contains one worksheet default.
    	workbook = new Workbook();

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
