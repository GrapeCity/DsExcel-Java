package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddCustomName extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet1 = workbook.getWorksheets().get(0);
        IWorksheet worksheet2 = workbook.getWorksheets().add();

        worksheet1.getRange("C8").setNumberFormat("0.0000");

        worksheet1.getNames().add("test1", "=Sheet1!$A$1");
        worksheet1.getNames().add("test2", "=Sheet1!test1*2");
        workbook.getNames().add("test3", "=Sheet1!$A$1");

        worksheet1.getRange("A1").setValue(1);

        // C6's value is 1.
        worksheet1.getRange("C6").setFormula("=test1");
        // C7's value is 3.
        worksheet1.getRange("C7").setFormula("=test1 + test2");
        // C8's value is 6.283185307
        worksheet1.getRange("C8").setFormula("=test2*PI()");

        // judge if Range C6:C8 have formula.
        for (int i = 5; i <= 7; i++) {
            if (worksheet1.getRange(i, 2).getHasFormula()) {
                worksheet1.getRange(i, 2).getInterior().setColor(Color.GetLightBlue());
            }
        }

        // worksheet1 range A2's value is 1.
        worksheet2.getRange("A2").setFormula("=test3");
        // judge if Range A2 has formula.
        if (worksheet2.getRange("A2").getHasFormula()) {
            worksheet2.getRange("A2").getInterior().setColor(Color.GetLightBlue());
        }

        // set r1c1 formula.
        worksheet2.getRange("A3").setFormulaR1C1("=R[-1]C");

    }

}
