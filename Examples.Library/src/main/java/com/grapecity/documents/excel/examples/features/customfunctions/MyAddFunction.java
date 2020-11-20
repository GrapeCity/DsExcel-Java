package com.grapecity.documents.excel.examples.features.customfunctions;

import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MyAddFunction extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MyAddFunctionX());
        IWorksheet worksheet = workbook.getActiveSheet();
        worksheet.getRange("A1").setValue(1);
        worksheet.getRange("B1").setValue(2);
        worksheet.getRange("C1").setFormula("=MyAdd(A1, B1)");

        // Range C1's value is 3
        Object result = worksheet.getRange("C1").getValue();

        worksheet.getRange("E1:F2").setValue(new Object[][]{
            {1, 3},
            {2, 4}
        });

        worksheet.getRange("G1:G2").setFormulaArray("=MyAdd(E1:E2, F1:F2)");

        //Range G1's value is 4, Range G2's value is 6.
        Object resultG1 = worksheet.getRange("G1").getValue();
        Object resultG2 = worksheet.getRange("G2").getValue();

         /* Implementation of MyAddFunctionX

            class MyAddFunctionX extends CustomFunction {
                public MyAddFunctionX() {
                    super("MyAdd", FunctionValueType.Number, new Parameter[]{new Parameter(FunctionValueType.Number), new Parameter(FunctionValueType.Number)});
                }

                @Override
                public Object evaluate(Object[] arguments, ICalcContext context) {
                    return (double) arguments[0] + (double) arguments[1];
                }
            }

          */
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

}

class MyAddFunctionX extends CustomFunction {
    public MyAddFunctionX() {
        super("MyAdd", FunctionValueType.Number, new Parameter[]{new Parameter(FunctionValueType.Number), new Parameter(FunctionValueType.Number)});
    }

    @Override
    public Object evaluate(Object[] arguments, ICalcContext context) {
        return (double) arguments[0] + (double) arguments[1];
    }
}
