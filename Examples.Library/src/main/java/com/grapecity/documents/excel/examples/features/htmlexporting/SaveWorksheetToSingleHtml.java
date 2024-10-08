package com.grapecity.documents.excel.examples.features.htmlexporting;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.HtmlSaveOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class SaveWorksheetToSingleHtml extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/BreakEven.xlsx");
        workbook.open(fileStream);

        HtmlSaveOptions options = new HtmlSaveOptions();

        // Export first sheet
        options.setExportSheetName(workbook.getWorksheets().get(0).getName());

        // Set exported image as base64
        options.setExportImageAsBase64(true);

        // Set exported css style in html file
        options.setExportCssSeparately(false);

        // Set not to export single tab in html
        options.setExportSingleTab(false);

        workbook.save(outputStream, options);
    }

    @Override
    public boolean getSaveAsHtml() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/BreakEven.xlsx" };
    }
}
