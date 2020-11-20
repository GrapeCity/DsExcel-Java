package com.grapecity.documents.excel.examples.features.htmlexporting;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.HtmlSaveOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class SaveHtmlWithHeadingsAndGridlines extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Home maintenance schedule and task list.xlsx");
        workbook.open(fileStream);

        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set exporting row/column headings
        options.setExportHeadings(true);

        // Set exporting gridlines
        options.setExportGridlines(true);

        // Set exported html file name
        options.setExportFileName("schedule");

        workbook.save(outputStream, options);
    }

    @Override
    public boolean getSaveAsZip() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Home maintenance schedule and task list.xlsx" };
    }
}
