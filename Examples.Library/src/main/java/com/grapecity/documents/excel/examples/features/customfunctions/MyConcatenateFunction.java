package com.grapecity.documents.excel.examples.features.customfunctions;

import com.grapecity.documents.excel.CalcError;
import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MyConcatenateFunction extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MyConcatenateFunctionX());

        IWorksheet worksheet = workbook.getActiveSheet();
        worksheet.getRange("A1").setFormula("=MyConcatenate(\"I\", \" \", \"live\", \" \", \"in\", \" \", \"Xi'an\", \".\")");
        worksheet.getRange("A2").setFormula("=MyConcatenate(A1, \"haha.\")");

        worksheet.getRange("B1").setValue(12);
        worksheet.getRange("B2").setValue(34);
        worksheet.getRange("B3").setFormula("=MyConcatenate(B1, B2)");

        worksheet.getRange("M5:N5").setFormulaArray("=CONCATENATE({\"aa\",\"bb\"}, 12, 34)");

        //"I live in Xi'an."
        Object resultA1 = worksheet.getRange("A1").getValue();
        //"I live in Xi'an.haha."
        Object resultA2 = worksheet.getRange("A2").getValue();
        //"12.034.0"
        Object resultB3 = worksheet.getRange("B3").getValue();
        //"aa1234"
        Object resultM5 = worksheet.getRange("M5").getValue();
        //"bb1234"
        Object resultN5 = worksheet.getRange("N5").getValue();


         /* Implementation of MyAddFunctionX

            class MyConcatenateFunctionX extends CustomFunction {
                public MyConcatenateFunctionX() {
                    super("MyConcatenate", FunctionValueType.Text, CreateParameters());
                }

                static Parameter[] CreateParameters() {
                    Parameter[] parameters = new Parameter[254];
                    for (int i = 0; i < 254; i++) {
                        parameters[i] = new Parameter(FunctionValueType.Variant);
                    }

                    return parameters;
                }

                @Override
                public Object evaluate(Object[] arguments, ICalcContext context) {
                    StringBuilder sb = new StringBuilder();

                    for (Object argument : arguments) {
                        if (argument instanceof CalcError) {
                            return argument;
                        }
                        if (argument instanceof String || argument instanceof Double) {
                            sb.append(argument);
                        }
                    }

                    return sb.toString();
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


class MyConcatenateFunctionX extends CustomFunction {
    public MyConcatenateFunctionX() {
        super("MyConcatenate", FunctionValueType.Text, CreateParameters());
    }

    static Parameter[] CreateParameters() {
        Parameter[] parameters = new Parameter[254];
        for (int i = 0; i < 254; i++) {
            parameters[i] = new Parameter(FunctionValueType.Variant);
        }

        return parameters;
    }

    @Override
    public Object evaluate(Object[] arguments, ICalcContext context) {
        StringBuilder sb = new StringBuilder();

        for (Object argument : arguments) {
            if (argument instanceof CalcError) {
                return argument;
            }
            if (argument instanceof String || argument instanceof Double) {
                sb.append(argument);
            }
        }

        return sb.toString();
    }
}
