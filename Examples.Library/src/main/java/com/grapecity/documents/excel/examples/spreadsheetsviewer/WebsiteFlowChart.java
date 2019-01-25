package com.grapecity.documents.excel.examples.spreadsheetsviewer;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class WebsiteFlowChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getTemplateStream());
    }

    @Override
    public String getTemplateName() {
        return "WebsiteFlowChart.xlsx";
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
    public boolean getIsNew() {
        return true;
    }
}
