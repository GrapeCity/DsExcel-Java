package com.grapecity.documents.excel.examples.features.cellrichtext;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigRunFont extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange a2 = worksheet.getRange("A2");

        a2.getFont().setSize(18);
        a2.getFont().setBold(true);
        a2.setVerticalAlignment(VerticalAlignment.Center);

        a2.getEntireRow().setRowHeight(42);
        a2.getEntireColumn().setColumnWidth(40);
        a2.setValue("Perfect square trinomial");

        ITextRun run = a2.characters(8, 7);
        run.getFont().setItalic(true);
        run.getFont().setThemeColor(ThemeColor.Accent1);

        IRange b2 = worksheet.getRange("B2");
        b2.getFont().setSize(26);
        b2.getEntireColumn().setColumnWidth(60);

        b2.setValue("(a+b)2 = a2+2ab+b2");

        ITextRun superRun1 = b2.characters(5, 1);
        superRun1.getFont().setSuperscript(true);
        superRun1.getFont().setColor(Color.GetRed());

        ITextRun superRun2 = b2.characters(10, 1);
        superRun2.getFont().setSuperscript(true);
        superRun2.getFont().setColor(Color.GetGreen());

        ITextRun superRun3 = b2.characters(17, 1);
        superRun3.getFont().setSuperscript(true);
        superRun3.getFont().setColor(Color.GetBlue());
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
