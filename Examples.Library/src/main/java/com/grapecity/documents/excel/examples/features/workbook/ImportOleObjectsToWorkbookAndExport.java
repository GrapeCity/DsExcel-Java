package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportOleObjectsToWorkbookAndExport extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        workbook.open(getResourceStream("xlsx/OleTemplates.xlsx"));
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getCanDownload() {
        return true;
    }

    private String[] UsedResources = new String[] {"xlsx/OleTemplates.xlsx"};

    @Override
    public String[] getResources() {
        return UsedResources;
    }
}
