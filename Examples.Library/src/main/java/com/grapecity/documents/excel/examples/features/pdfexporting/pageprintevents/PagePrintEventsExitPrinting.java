package com.grapecity.documents.excel.examples.features.pdfexporting.pageprintevents;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PagePrintEventsExitPrinting extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        IWorksheet activeSheet = workbook.getActiveSheet();
        activeSheet.getRange("A1").setValue(1);
        activeSheet.getRange("A2:A100").setFormulaR1C1("=R[-1]C+1");
        PdfSaveOptions options = new PdfSaveOptions();
        options.getPagePrintedEvent().addListener((sender, e) -> {
            if (e.getPageNumber() == 2) {
                e.setHasMorePages(false);
            }
        });
        activeSheet.getPageSetup().setCenterHeader("Page &P of &N");
        workbook.save(outputStream, options);
    }

    @Override
    public boolean getSavePageInfos() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
    @Override
    public boolean getShowViewer() {
        return false;
    }
}