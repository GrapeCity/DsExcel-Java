package com.grapecity.documents.excel.examples.features.formatting;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class TextEllipsis extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        InputStream fileStream = this.getResourceStream("json/TextEllipsis.json");
        workbook.fromJson(fileStream);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "json/TextEllipsis.json" };
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
