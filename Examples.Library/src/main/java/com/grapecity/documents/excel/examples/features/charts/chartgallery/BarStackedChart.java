package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LegendPosition;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BarStackedChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.BarStacked, 250, 20, 360, 230);
        worksheet.getRange("A1:C4").setValue(new Object[][]{
                {103, 121, 109},
                {56, 94, 115},
                {116, 89, 99},
                {55, 93, 70}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C4"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Bar Stacked Chart");
        shape.getChart().getLegend().setPosition(LegendPosition.Left);
    }

}
