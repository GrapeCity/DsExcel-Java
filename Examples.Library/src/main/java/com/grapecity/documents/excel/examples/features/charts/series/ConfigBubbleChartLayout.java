package com.grapecity.documents.excel.examples.features.charts.series;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.SizeRepresents;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigBubbleChartLayout extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().addChart(ChartType.Bubble, 250, 20, 360, 230);
        worksheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        ISeries series1 = shape.getChart().getSeriesCollection().get(0);
        series1.setHasDataLabels(true);

        shape.getChart().getXYGroups().get(0).setBubbleScale(150);
        shape.getChart().getXYGroups().get(0).setSizeRepresents(SizeRepresents.SizeIsArea);
        shape.getChart().getXYGroups().get(0).setShowNegativeBubbles(true);

    }

}
