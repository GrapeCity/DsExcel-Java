package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AreaStackedChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.AreaStacked, 250, 20, 360, 230);
        worksheet.getRange("A1:C13").setValue(new Object[][]{
                {0, 59.18, 27.14},
                {44.64, 52.22, 25.08},
                {45.21, 49.80, 57.99},
                {24.32, 37.30, 42.73},
                {58.34, 34.43, 28.34},
                {31.89, 69.78, 46.88},
                {41.79, 63.94, 56.24},
                {67.94, 57.40, 27.78},
                {49.87, 48.26, 52.06},
                {62.39, 67.43, 33.33},
                {54.76, 22.95, 50.36},
                {28.33, 36.60, 36.61},
                {22.77, 55.65, 65.64},
                {20.34, 49.35, 45.60},
                {32.10, 47.60, 20.62},
                {26.37, 63.00, 53.97},
                {35, 75, 60}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C13"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Area Stacked Chart");
    }

}
