package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LegendPosition;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PieChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Pie, 250, 20, 360, 230);
        worksheet.getRange("A1:B4").setValue(new Object[][]{
                {"Blue", 1},
                {"Red", 2},
                {"Green", 3},
                {"Purple", 4},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B4"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Pie Chart");
        shape.getChart().getLegend().setPosition(LegendPosition.Right);

    }

}
