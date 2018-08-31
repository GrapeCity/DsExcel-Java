package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BubbleChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Bubble, 250, 20, 360, 230);
        worksheet.getRange("A1:C10").setValue(new Object[][]{
                {"Blue", null, null},
                {125, 750, 3},
                {25, 625, 7},
                {75, 875, 5},
                {175, 625, 6},
                {"Red", null, null},
                {125, 500, 10},
                {25, 250, 1},
                {75, 125, 5},
                {175, 250, 8},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C5"), RowCol.Columns);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A6:C10"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Bubble Chart");
    }

}
