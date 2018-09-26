package com.grapecity.documents.excel.examples.features.customfunctions;

import com.grapecity.documents.excel.CalcError;
import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MyIsErrorFunction extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MyIsErrorFunctionX());

        IWorksheet worksheet = workbook.getActiveSheet();

        worksheet.getRange("A1").setValue(CalcError.Num);
        worksheet.getRange("A2").setValue(100);

        worksheet.getRange("B1").setFormula("=MyIsError(A1)");
        worksheet.getRange("B2").setFormula("=MyIsError(A2)");

        //getRange("B1")'s value is true.
        Object resultB1 = worksheet.getRange("B1").getValue();

        //getRange("B2")'s value is false.
        Object resultB2 = worksheet.getRange("B2").getValue();

         /* Implementation of MyAddFunctionX

            class MyIsErrorFunctionX extends CustomFunction {
                public MyIsErrorFunctionX() {
                    super("MyIsError", FunctionValueType.Boolean, new Parameter[]{new Parameter(FunctionValueType.Variant)});
                }

                @Override
                public Object evaluate(Object[] arguments, ICalcContext context) {
                    if (arguments[0] instanceof CalcError) {
                        if ((CalcError) arguments[0] != CalcError.None && (CalcError) arguments[0] != CalcError.GettingData) {
                            return true;
                        } else {
                            return false;
                        }
                    }

                    return false;
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


class MyIsErrorFunctionX extends CustomFunction {
    public MyIsErrorFunctionX() {
        super("MyIsError", FunctionValueType.Boolean, new Parameter[]{new Parameter(FunctionValueType.Variant)});
    }

    @Override
    public Object evaluate(Object[] arguments, ICalcContext context) {
        if (arguments[0] instanceof CalcError) {
            if ((CalcError) arguments[0] != CalcError.None && (CalcError) arguments[0] != CalcError.GettingData) {
                return true;
            } else {
                return false;
            }
        }

        return false;
    }
}
