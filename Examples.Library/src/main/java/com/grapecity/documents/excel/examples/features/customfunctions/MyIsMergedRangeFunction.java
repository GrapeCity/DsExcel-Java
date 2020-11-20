package com.grapecity.documents.excel.examples.features.customfunctions;

import java.util.List;

import com.grapecity.documents.excel.CalcReference;
import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MyIsMergedRangeFunction extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MyIsMergedRangeFunctionX());

        IWorksheet worksheet = workbook.getActiveSheet();

        worksheet.getRange("A1:B2").merge();

        worksheet.getRange("C1").setFormula("=MyIsMergedRange(A1)");
        worksheet.getRange("C2").setFormula("=MyIsMergedRange(H2)");

        //A1 is a merged cell, getRange("C1")'s value is true.
        Object resultC1 = worksheet.getRange("C1").getValue();

        //H2 is not a merged cell, getRange("C2")'s value is false.
        Object resultC2 = worksheet.getRange("C2").getValue();

         /* Implementation of MyAddFunctionX

            class MyIsMergedRangeFunctionX extends CustomFunction {
                public MyIsMergedRangeFunctionX() {
                    super("MyIsMergedRange", FunctionValueType.Boolean, new Parameter[]{new Parameter(FunctionValueType.Object, true)});
                }

                @Override
                public Object evaluate(Object[] arguments, ICalcContext context) {
                    if (arguments[0] instanceof CalcReference) {
                        if (arguments[0] instanceof CalcReference) {
                            List<IRange> ranges = ((CalcReference) arguments[0]).getRanges();
                            for (IRange range : ranges) {
                                return range.getMergeCells();
                            }
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


class MyIsMergedRangeFunctionX extends CustomFunction {
    public MyIsMergedRangeFunctionX() {
        super("MyIsMergedRange", FunctionValueType.Boolean, new Parameter[]{new Parameter(FunctionValueType.Object, true)});
    }

    @Override
    public Object evaluate(Object[] arguments, ICalcContext context) {
        if (arguments[0] instanceof CalcReference) {
            if (arguments[0] instanceof CalcReference) {
                List<IRange> ranges = ((CalcReference) arguments[0]).getRanges();
                for (IRange range : ranges) {
                    return range.getMergeCells();
                }
            }
        }

        return false;
    }
}
