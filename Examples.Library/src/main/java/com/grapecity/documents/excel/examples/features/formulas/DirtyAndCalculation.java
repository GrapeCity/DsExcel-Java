package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DirtyAndCalculation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue(1);
        worksheet.getRange("A2").setFormula("=A1");
        worksheet.getRange("A3").setFormula("=SUM(A1, A2)");

        //when get value, calc engine will first calculate and cache the result, then returns the cached result.
        Object value_A2 = worksheet.getRange("A2").getValue();
        Object value_A3 = worksheet.getRange("A3").getValue();

        //disable calc engine.
        workbook.setEnableCalculation(false);

        //Dirty() method will clear the cached value of the workbook.
        workbook.dirty();
        //Calculate() will not work, because of workbook.EnablCalculation is false.
        workbook.calculate();
        //it returns 0 because of no cache value exist.
        Object value_A2_1 = worksheet.getRange("A2").getValue();
        Object value_A3_1 = worksheet.getRange("A3").getValue();

        worksheet.getRange("A1").setValue(2);
        //enable calc engine.
        workbook.setEnableCalculation(true);
        //Dirty() method will clear the cached value of Range A2:A3.
        //    worksheet.get("A2:A3").Dirty();
        //Calculate() method will calculate and cache the result, it will return the cache value directly when get value later.
        //   worksheet.get("A2:A3").Calculate();

        //it returns cache value directly, does not calculate again.
        Object value_A2_2 = worksheet.getRange("A2").getValue();
        Object value_A3_2 = worksheet.getRange("A3").getValue();

    }

}
