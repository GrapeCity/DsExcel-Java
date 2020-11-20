package com.grapecity.documents.excel.examples.features.spreadjsfeatures;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class CellButtons extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        InputStream fileStream = this.getResourceStream("json/CellButtons.json");
        workbook.fromJson(fileStream);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "json/CellButtons.json" };
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
