package com.grapecity.documents.excel.examples.features.datavalidation;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class HighlightStyle extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        InputStream fileStream = this.getResourceStream("json/HighlightStyle.json");
        workbook.fromJson(fileStream);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "json/HighlightStyle.json" };
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
