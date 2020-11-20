package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SelectWorksheets extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet1 = workbook.getActiveSheet();
        IWorksheet sheet2 = workbook.getWorksheets().add();
        IWorksheet sheet3 = workbook.getWorksheets().add();

        // Select sheet2 and sheet3.
        workbook.getWorksheets().get(new String[] { sheet2.getName(), sheet3.getName() }).select();

        // Write names of selected sheets to console
        for (IWorksheet sheet : workbook.getSelectedSheets()) {
            System.out.println(sheet.getName());
        }

        // Add sheet1 to selected sheets
        sheet1.select(false);

        // Write count of selected sheets to console
        System.out.println(workbook.getSelectedSheets().getCount());
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

}
