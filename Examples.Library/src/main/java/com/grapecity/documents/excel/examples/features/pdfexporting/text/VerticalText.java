package com.grapecity.documents.excel.examples.features.pdfexporting.text;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class VerticalText extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IWorksheet sheet = workbook.getWorksheets().get(0);

        IRange a1 = sheet.getRange("A1");
        a1.getFont().setName("@Meiryo");
        a1.setValue("日本語（にほんご、にっぽんご）は、主として、日本列島で使用されてきた言語である。");
        a1.setHorizontalAlignment(HorizontalAlignment.Right);
        a1.setVerticalAlignment(VerticalAlignment.Top);
        a1.setOrientation(-90);
        a1.setWrapText(true);

        a1.setColumnWidth(27);
        a1.setRowHeight(190);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}