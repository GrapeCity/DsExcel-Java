package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddWorksheet extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        // Add a new worksheet to the workbook. The worksheet will be inserted into the end of the existing worksheet collection.
        workbook.getWorksheets().add();

        //Add a new worksheet to the specified position in the collection of worksheets.
        workbook.getWorksheets().addBefore(workbook.getWorksheets().get(0));
        workbook.getWorksheets().addAfter(workbook.getWorksheets().get(1));

        //Set worksheet's name.
        workbook.getWorksheets().get(2).setName("Product Plan");

    }

}
