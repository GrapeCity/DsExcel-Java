package com.grapecity.documents.excel.examples.features.charts.axes;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.ScaleType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetAxisScaleType extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D5").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 4, 25, 7},
                {"Item2", 15, -10, 18},
                {"Item3", 45, 90, 20},
                {"Item4", 8, 20, 11}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D5"), RowCol.Columns, true, true);

        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value);
        value_axis.setScaleType(ScaleType.Logarithmic);
        value_axis.setLogBase(5);

    }

}
