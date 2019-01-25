package com.grapecity.documents.excel.examples.features.cellrichtext;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CharacterCellText extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange b2 = worksheet.getRange("B2");
        b2.setValue("GrapeCity Documents for Excel");
        b2.getFont().setSize(26);
        b2.getEntireRow().setRowHeight(42);

        ITextRun run1 = b2.characters(0, 9);
        run1.getFont().setName("Agency FB");
        run1.getFont().setThemeColor(ThemeColor.Accent1);
        run1.getFont().setBold(true);

        ITextRun run2 = b2.characters(10, 9);
        run2.getFont().setThemeColor(ThemeColor.Accent2);
        run2.getFont().setName("Arial Black");
        run2.getFont().setUnderline(UnderlineType.Single);

        ITextRun run3 = b2.characters(20, 3);
        run3.getFont().setItalic(true);

        ITextRun run4 = b2.characters(24, 5);
        run4.getFont().setColor(Color.GetBlue());
        run4.getFont().setBold(true);

    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
