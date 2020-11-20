package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class LineMarkerStacked100Chart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.LineMarkersStacked100, 250, 20, 360, 230);
        worksheet.getRange("A1:C5").setValue(new Object[][]{
                {12, 22, 27},
                {45, 52, 25},
                {58, 35, 58},
                {21, 37, 43},
                {44, 45, 28}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C5"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Line Marker Stacked 100 Chart");

    }

}
