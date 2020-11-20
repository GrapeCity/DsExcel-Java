package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class XYScatterSmooth extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.XYScatterSmoothNoMarkers, 350, 20, 360, 230);
        worksheet.getRange("A1:B5").setValue(new Object[][]{
                {4, 2},
                {6, 1},
                {1, 2},
                {7, 4},
                {4, 4},
        });
        worksheet.getRange("A7:B11").setValue(new Object[][]{
                {9, 5},
                {7, 8},
                {9, 8},
                {5, 9},
                {2, 4},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B5"), RowCol.Columns);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A7:B11"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Scatter with Smooth Lines Chart");

    }

}
