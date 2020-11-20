package com.grapecity.documents.excel.examples.features.formatting;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class FormatString extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        InputStream fileStream = this.getResourceStream("json/FormatString.ssjson");
        workbook.fromJson(fileStream);
    }

    @Override
    public boolean getSaveAsJSON() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "json/FormatString.ssjson" };
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
