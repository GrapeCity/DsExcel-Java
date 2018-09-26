package com.grapecity.documents.excel.examples.features.formulas;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class UseArrayFormula extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("E4:J5").setValue(new Object[][]{
                {1, 2, 3},
                {4, 5, 6}
        });

        worksheet.getRange("I6:J8").setValue(new Object[][]{

                {2, 2},
                {3, 3},
                {4, 4}
        });

        //O     P      Q
        //2     4      #N/A
        //12    15     #N/A
        //#N/A  #N/A   #N/A
        worksheet.getRange("O9:Q11").setFormulaArray("=E4:G5*I6:J8");

        //judge if Range O9 has array formula.
        if (worksheet.getRange("O9").getHasArray()) {
            //set O9's entire array's interior color.
            IRange currentarray = worksheet.getRange("O9").getCurrentArray();
            currentarray.getInterior().setColor(Color.GetGreen());
        }
    }

}
