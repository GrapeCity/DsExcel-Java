package com.grapecity.documents.excel.examples.features.customfunctions;

import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.CalcError;
import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MySumFunction extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MySumFunctionX());

        IWorksheet worksheet = workbook.getActiveSheet();
        worksheet.getRange("A1").setValue(1);
        worksheet.getRange("B1").setValue(2);
        worksheet.getRange("C1").setFormula("=MySum(A1:B1, 2, {3,4})");

        //getRange("C1")'s value is 12.
        Object result = worksheet.getRange("C1").getValue();


         /* Implementation of MyAddFunctionX

            class MySumFunctionX extends CustomFunction {
                public MySumFunctionX() {
                    super("MySum", FunctionValueType.Number, CreateParameters());
                }

                private static Parameter[] CreateParameters() {
                    Parameter[] parameters = new Parameter[254];
                    for (int i = 0; i < 254; i++) {
                        parameters[i] = new Parameter(FunctionValueType.Object);
                    }

                    return parameters;
                }

                @Override
                public Object evaluate(Object[] arguments, ICalcContext context) {
                    double sum = 0d;
                    for (Object argument : arguments) {
                        Iterable<Object> iterator = toIterable(argument);
                        for (Object item : iterator) {
                            if (item instanceof CalcError) {
                                return item;
                            }
                            if (item instanceof Double) {
                                sum += (double) item;
                            }
                        }
                    }

                    return sum;
                }

                private static Iterable<Object> toIterable(Object obj) {
                    if (obj instanceof Iterable) {
                        return (Iterable) obj;
                    } else if (obj instanceof Object[][]) {
                        List<Object> list = new ArrayList<Object>();
                        Object[][] array = (Object[][]) obj;
                        for (int i = 0; i < array.length; i++) {
                            for (int j = 0; j < array[i].length; j++) {
                                list.add(array[i][j]);
                            }
                        }
                        return list;
                    } else {
                        List<Object> list = new ArrayList<Object>();
                        list.add(obj);
                        return list;
                    }
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


class MySumFunctionX extends CustomFunction {
    public MySumFunctionX() {
        super("MySum", FunctionValueType.Number, CreateParameters());
    }

    private static Parameter[] CreateParameters() {
        Parameter[] parameters = new Parameter[254];
        for (int i = 0; i < 254; i++) {
            parameters[i] = new Parameter(FunctionValueType.Object);
        }

        return parameters;
    }

    @Override
    public Object evaluate(Object[] arguments, ICalcContext context) {
        double sum = 0d;
        for (Object argument : arguments) {
            Iterable<Object> iterator = toIterable(argument);
            for (Object item : iterator) {
                if (item instanceof CalcError) {
                    return item;
                }
                if (item instanceof Double) {
                    sum += (double) item;
                }
            }
        }

        return sum;
    }

    private static Iterable<Object> toIterable(Object obj) {
        if (obj instanceof Iterable) {
            return (Iterable) obj;
        } else if (obj instanceof Object[][]) {
            List<Object> list = new ArrayList<Object>();
            Object[][] array = (Object[][]) obj;
            for (int i = 0; i < array.length; i++) {
                for (int j = 0; j < array[i].length; j++) {
                    list.add(array[i][j]);
                }
            }
            return list;
        } else {
            List<Object> list = new ArrayList<Object>();
            list.add(obj);
            return list;
        }
    }
}
