package com.grapecity.documents.excel.examples.features.pdfexporting.text;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ExportCellRichText extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange a1 = worksheet.getRange("A1");

        a1.getFont().setSize(18);
        a1.getFont().setBold(true);
        a1.setVerticalAlignment(VerticalAlignment.Center);

        a1.getEntireRow().setRowHeight(42);
        a1.getEntireColumn().setColumnWidth(35);
        a1.setValue("Perfect square trinomial");

        ITextRun run = a1.characters(8, 7);
        run.getFont().setItalic(true);
        run.getFont().setThemeColor(ThemeColor.Accent1);

        IRange b1 = worksheet.getRange("B1");
        b1.getFont().setSize(26);
        b1.getEntireColumn().setColumnWidth(40);

        b1.setValue("(a+b)2 = a2+2ab+b2");

        ITextRun superRun1 = b1.characters(5, 1);
        superRun1.getFont().setSuperscript(true);
        superRun1.getFont().setColor(Color.GetRed());

        ITextRun superRun2 = b1.characters(10, 1);
        superRun2.getFont().setSuperscript(true);
        superRun2.getFont().setColor(Color.GetGreen());

        ITextRun superRun3 = b1.characters(17, 1);
        superRun3.getFont().setSuperscript(true);
        superRun3.getFont().setColor(Color.GetBlue());
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