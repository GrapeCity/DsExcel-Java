package com.grapecity.documents.excel.examples.features.charts.errorbars;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigXYErrorBars extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.XYScatter, 250, 20, 360, 230);
        worksheet.getRange("A1:D7").setValue(new Object[][]{
                {"Blue", null, "Red", null},
                {55, 964, 67, 475},
                {20, 825, 10, 163},
                {77, 840, 87, 224},
                {182, 596, 46, 196},
                {190, 384, 100, 377},
                {140, 503, 92, 47},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A2:B7"), RowCol.Columns);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("C2:D7"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("Scatter Chart");

        // Get first series
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);

        // Set HasErrorBars as true
        series1.setHasErrorBars(true);

        // Config y-direction error bar
        series1.getYErrorBar().setValueType(ErrorBarType.FixedValue);
        series1.getYErrorBar().setAmount(500);

        // Config x-direction error bar
        series1.getXErrorBar().setValueType(ErrorBarType.FixedValue);
        series1.getXErrorBar().setAmount(20);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
