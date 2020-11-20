package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ColumnStackedChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnStacked, 250, 20, 360, 230);
        worksheet.getRange("A1:C6").setValue(new Object[][]{
                {103, 121, 109},
                {56, 94, 115},
                {116, 89, 99},
                {55, 93, 70},
                {114, 114, 83},
                {125, 138, 136}
        });

        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C6"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Column Stacked Chart");

    }

}
