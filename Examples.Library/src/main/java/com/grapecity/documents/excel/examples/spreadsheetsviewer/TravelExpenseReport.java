package com.grapecity.documents.excel.examples.spreadsheetsviewer;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class TravelExpenseReport extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getTemplateStream());
    }

    @Override
    public String getTemplateName() {
        return "Travel expense report.xlsx";
    }

    @Override
    public boolean getHasTemplate() {
        return true;
    }

    @Override
    public boolean getIsViewReadOnly() {
        return false;
    }

    @Override
    public boolean getShowCode() {
        return false;
    }

    @Override
    public boolean getCanDownloadZip() {
        return false;
    }
}
