package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CombinationChart1 extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        worksheet.getRange("A1:B13").setValue(new Object[][]{
                {"Blue Column", "Red Line"},
                {75, 20},
                {149, 50},
                {105, 30},
                {55, 80},
                {121, 40},
                {76, 110},
                {128, 50},
                {114, 140},
                {75, 60},
                {105, 170},
                {145, 70},
                {110, 100}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B13"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Combination Chart");
        //change series type
        ISeries series2 = shape.getChart().getSeriesCollection().get(1);
        series2.setChartType(ChartType.LineMarkers);

    }

}
