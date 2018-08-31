package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LegendPosition;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BarStacked100Chart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.BarStacked100, 250, 20, 360, 230);
        worksheet.getRange("A1:B5").setValue(new Object[][]{
                {1, 5},
                {2, 4},
                {3, 3},
                {4, 2},
                {4, 1}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B5"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Bar Stacked 100 Chart");
        shape.getChart().getLegend().setPosition(LegendPosition.Left);
    }

}
