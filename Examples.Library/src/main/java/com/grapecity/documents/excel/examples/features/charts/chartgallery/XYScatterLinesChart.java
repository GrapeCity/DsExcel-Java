package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class XYScatterLinesChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.XYScatterLines, 350, 20, 360, 230);
        worksheet.getRange("A1:B8").setValue(new Object[][]{
                {75, 250},
                {50, 125},
                {25, 375},
                {75, 250},
                {50, 875},
                {25, 625},
                {75, 750},
                {125, 500},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B8"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Scatter with Straight Lines and Markers Chart");

    }

}
