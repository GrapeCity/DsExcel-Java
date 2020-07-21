package com.grapecity.documents.excel.examples.features.htmlexporting;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.HtmlSaveOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class SaveHtmlWithScalableWidth extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/NetProfit.xlsx");
        workbook.open(fileStream);

        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set html with scalable width
        options.setIsWidthScalable(true);

        // Set exported html file name
        options.setExportFileName("NetProfit");

        workbook.save(outputStream, options);
    }

    @Override
    public boolean getSaveAsZip() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/NetProfit.xlsx" };
    }
}
