package com.grapecity.documents.excel.examples.features.customfunctions;

import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.CalcError;
import com.grapecity.documents.excel.CalcReference;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.CustomFunction;
import com.grapecity.documents.excel.FormatConditionOperator;
import com.grapecity.documents.excel.FormatConditionType;
import com.grapecity.documents.excel.FunctionValueType;
import com.grapecity.documents.excel.ICalcContext;
import com.grapecity.documents.excel.IFormatCondition;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Parameter;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MyConditionalSumFunction extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook.AddCustomFunction(new MyConditionalSumFunctionX());
        IWorksheet worksheet = workbook.getActiveSheet();
        worksheet.getRange("A1:A10").setValue(new Object[][] {
            {1 },
            {2 },
            {3 },
            {4 },
            {5 },
            {6 },
            {7 },
            {8 },
            {9 },
            {10 }});
        IFormatCondition cellValueRule = (IFormatCondition)worksheet.getRange("A1:A10").getFormatConditions().add(FormatConditionType.CellValue, FormatConditionOperator.Greater, 5, null);
        cellValueRule.getInterior().setColor(Color.GetRed());

        //Sum cells value which display format interior color are red.
        worksheet.getRange("C1").setFormula("=MyConditionalSum(A1:A10)");

        //Range["C1"]'s value is 40.
        Object result = worksheet.getRange("C1").getValue();


         /* Implementation of MyAddFunctionX

              class MyConditionalSumFunctionX extends CustomFunction {
                public MyConditionalSumFunctionX() {
                    super("MyConditionalSum", FunctionValueType.Number, CreateParameters());
                }

                private static Parameter[] CreateParameters() {
                    Parameter[] parameters = new Parameter[254];
                    for (int i = 0; i < 254; i++) {
                        parameters[i] = new Parameter(FunctionValueType.Object, true);
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
                            } else if (item instanceof Double) {
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
                    } else if (obj instanceof CalcReference) {
                        List<Object> list = new ArrayList<Object>();
                        CalcReference reference = (CalcReference) obj;
                        for (IRange range : reference.getRanges()) {
                            int rowCount = range.getRows().getCount();
                            int colCount = range.getColumns().getCount();
                            for (int i = 0; i < rowCount; i++) {
                                for (int j = 0; j < colCount; j++) {
                                    if (range.getCells().get(i, j).getDisplayFormat().getInterior().getColor().equals(Color.getRed())) {
                                        list.add(range.getCells().get(i, j).getValue());
                                    }
                                }
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


class MyConditionalSumFunctionX extends CustomFunction {
    public MyConditionalSumFunctionX() {
        super("MyConditionalSum", FunctionValueType.Number, CreateParameters());
    }

    private static Parameter[] CreateParameters() {
        Parameter[] parameters = new Parameter[254];
        for (int i = 0; i < 254; i++) {
            parameters[i] = new Parameter(FunctionValueType.Object, true);
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
                } else if (item instanceof Double) {
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
        } else if (obj instanceof CalcReference) {
            List<Object> list = new ArrayList<Object>();
            CalcReference reference = (CalcReference) obj;
            for (IRange range : reference.getRanges()) {
                int rowCount = range.getRows().getCount();
                int colCount = range.getColumns().getCount();
                for (int i = 0; i < rowCount; i++) {
                    for (int j = 0; j < colCount; j++) {
                        if (range.getCells().get(i, j).getDisplayFormat().getInterior().getColor().equals(Color.GetRed())) {
                            list.add(range.getCells().get(i, j).getValue());
                        }
                    }
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
