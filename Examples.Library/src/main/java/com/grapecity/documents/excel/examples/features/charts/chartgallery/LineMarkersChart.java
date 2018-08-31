package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.MarkerStyle;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class LineMarkersChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.LineMarkers, 250, 20, 360, 230);
        worksheet.getRange("A1:B8").setValue(new Object[][]{
                {6, 55},
                {45, 25},
                {35, 45},
                {25, 65},
                {65, 15},
                {45, 75},
                {75, 55},
                {65, 35},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B8"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Line with Markers");
        shape.getChart().getSeriesCollection().get(0).setMarkerStyle(MarkerStyle.Square);

    }

}
