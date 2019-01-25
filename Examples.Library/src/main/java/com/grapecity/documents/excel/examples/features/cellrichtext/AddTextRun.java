package com.grapecity.documents.excel.examples.features.cellrichtext;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddTextRun extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange b2 = worksheet.getRange("B2");

        ITextRun run1 = b2.getRichText().add("GrapeCity");
        run1.getFont().setName("Agency FB");
        run1.getFont().setSize(26);
        run1.getFont().setThemeColor(ThemeColor.Accent1);
        run1.getFont().setBold(true);

        ITextRun run2 = b2.getRichText().add(" Documents");
        run2.getFont().setThemeColor(ThemeColor.Accent2);
        run2.getFont().setName("Arial Black");
        run2.getFont().setSize(20);
        run2.getFont().setUnderline(UnderlineType.Single);

        ITextRun run3 = b2.getRichText().add(" for ");
        run3.getFont().setItalic(true);

        ITextRun run4 = b2.getRichText().add("Excel");
        run4.getFont().setColor(Color.GetBlue());
        run4.getFont().setBold(true);
        run4.getFont().setSize(26);
        run4.getFont().setUnderline(UnderlineType.Double);

        b2.getEntireRow().setRowHeight(42);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
