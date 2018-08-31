package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Pie_DoughnutChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Doughnut, 250, 20, 360, 230);
        worksheet.getRange("A1:B6").setValue(new Object[][]{
                {"S1", "S2"},
                {10, 25},
                {51, 36},
                {52, 85},
                {22, 65},
                {23, 69},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B6"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Area Chart");
        shape.getChart().getChartGroups().get(0).setDoughnutHoleSize(50);
        shape.getChart().getSeriesCollection().get(0).setHasDataLabels(true);
        shape.getChart().getSeriesCollection().get(1).setHasDataLabels(true);
        shape.getChart().getSeriesCollection().get(1).setExplosion(2);

    }

}
