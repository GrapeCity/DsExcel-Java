package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class UnprotectWorksheet extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //protect worksheet, allow insert column.
        worksheet.setProtection(true);
        worksheet.getProtectionSettings().setAllowInsertingColumns(true);

        //Unprotect worksheet.
        worksheet.setProtection(false);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
