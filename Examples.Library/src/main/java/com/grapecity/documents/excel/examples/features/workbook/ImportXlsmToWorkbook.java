package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.OpenFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportXlsmToWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        // GcExcel supports open xlsm file
        workbook.open("macros.xlsm");

        // Macros can be preserved after saving
        workbook.save("macros-exported.xlsm");
    }

    @Override
    public boolean getCanDownload() {
        return false;
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
