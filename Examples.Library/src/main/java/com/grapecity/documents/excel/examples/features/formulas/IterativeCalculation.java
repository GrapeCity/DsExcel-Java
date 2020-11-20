package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class IterativeCalculation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.getOptions().getFormulas().setEnableIterativeCalculation(true);
        workbook.getOptions().getFormulas().setMaximumIterations(20);
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // set values and formulas
        worksheet.getRange("B2").setValue("Initial Cash");
        worksheet.getRange("C2").setValue(10000);
        worksheet.getRange("B3").setValue("Interest");
        worksheet.getRange("C3").setValue(0.0125);

        worksheet.getRange("B5").setValue("Month");
        worksheet.getRange("C5").setValue("Total Cash");

        worksheet.getRange("B6:B26").setValue(new double[] {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21 });

        worksheet.getRange("C6").setFormula("=C2*(1+$C$3)");
        worksheet.getRange("C7:C26").setFormula("=C6*(1+$C$3)");

        worksheet.getRange("F2").setValue("Initial Cash");
        worksheet.getRange("G2").setValue(10000);

        worksheet.getRange("F3").setValue("Interest");
        worksheet.getRange("G3").setValue(0.0125);

        worksheet.getRange("F4").setValue("Total Cash at 21th Month");
        worksheet.getRange("G4").setFormula("=G2*(1+G3)");

        // this formula will generate circle reference.
        worksheet.getRange("G2").setFormula("=G4");

        // set styles
        worksheet.getRange("A:A,D:E").setColumnWidthInPixel(40);
        worksheet.getRange("B:C,F:G").setColumnWidthInPixel(100);

        worksheet.getRange("C2,G2,G4,C6:C26").setNumberFormat("$#,##0.00");
        worksheet.getRange("C3,G3").setNumberFormat("0.00%");

        worksheet.getRange("B2:C3,F2:G4,B5:C26").getBorders().setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("C2:C3,G2:G4").getInterior().setThemeColor(ThemeColor.Accent1);
        worksheet.getRange("C2:C3,G2:G4").getInterior().setTintAndShade(0.8);

        worksheet.getTables().add(worksheet.getRange("B5:C26"), true);

        // set print settings
        worksheet.getPageSetup().setPrintHeadings(true);
    }

    @Override
    public boolean getIsNew()
    {
        return true;
    }
}
