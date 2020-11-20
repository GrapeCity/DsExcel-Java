package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class LineChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 250, 20, 360, 230);
        worksheet.getRange("A1:C7").setValue(new Object[][]{
                {0, 59.18, 27.14},
                {44.64, 52.22, 25.08},
                {45.21, 49.80, 57.99},
                {24.32, 37.30, 42.73},
                {58.34, 34.43, 28.34},
                {31.89, 69.78, 46.88},
                {41.79, 63.94, 56.24}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C7"), RowCol.Columns);
        //set series lines style
        shape.getChart().getSeriesCollection().get(0).getFormat().getLine().setWeight(2.25);
        shape.getChart().getSeriesCollection().get(1).getFormat().getLine().setWeight(2.25);
        shape.getChart().getSeriesCollection().get(2).getFormat().getLine().setWeight(2.25);
        shape.getChart().getChartTitle().setText("Line Chart");

    }

}
