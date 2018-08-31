package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.MarkerStyle;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class XYScatterChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.XYScatter, 350, 20, 360, 230);
        worksheet.getRange("A1:D7").setValue(new Object[][]{
                {"Blue", null, "Red", null},
                {55, 964, 67, 475},
                {20, 825, 10, 163},
                {77, 840, 87, 224},
                {182, 596, 46, 196},
                {190, 384, 100, 377},
                {140, 503, 92, 47},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B7"), RowCol.Columns);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("C1:D7"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Scatter Chart");
        //config markers style
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);
        ISeries series2 = shape.getChart().getSeriesCollection().get(1);
        series1.setMarkerStyle(MarkerStyle.Square);
        series1.setMarkerSize(10);
        series2.setMarkerSize(10);

    }

}
