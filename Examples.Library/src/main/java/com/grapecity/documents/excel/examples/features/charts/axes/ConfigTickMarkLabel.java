package com.grapecity.documents.excel.examples.features.charts.axes;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.TickLabelPosition;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigTickMarkLabel extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);
        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value);
        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);

        //config tick label's format
        category_axis.setTickLabelPosition(TickLabelPosition.NextToAxis);
        category_axis.setTickLabelSpacing(2);
        category_axis.getTickLabels().getFont().getColor().setRGB(Color.GetDarkOrange());
        category_axis.getTickLabels().getFont().setSize(12);
        category_axis.getTickLabels().setNumberFormat("#,##0.00");
        value_axis.getTickLabels().setNumberFormat("#,##0;[Red]#,##0");

    }

}
