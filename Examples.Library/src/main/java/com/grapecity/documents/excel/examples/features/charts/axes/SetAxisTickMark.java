package com.grapecity.documents.excel.examples.features.charts.axes;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.chart.AxisType;
import com.grapecity.documents.excel.drawing.chart.ChartType;
import com.grapecity.documents.excel.drawing.chart.RowCol;
import com.grapecity.documents.excel.drawing.chart.TickMark;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.style.color.Color;

public class SetAxisTickMark extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);
        category_axis.getFormat().getLine().getColor().setRGB(Color.getGreen());
        category_axis.getFormat().getLine().setWeight(1.5);
        category_axis.setMajorTickMark(TickMark.Inside);
        category_axis.setMinorTickMark(TickMark.Cross);
        category_axis.setTickMarkSpacing(2);

    }

}
