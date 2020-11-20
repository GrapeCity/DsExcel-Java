package com.grapecity.documents.excel.examples.features.charts.errorbars;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigCustomErrorBar extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 250, 20, 360, 230);
        worksheet.getRange("A1:D4").setValue(new Object[][]{
                {null, "Q1", "Q2", "Q3"},
                {"Mobile Phones", 1330, 2330, 3330},
                {"Laptops", 4032, 5632, 6197},
                {"Tablets", 6233, 7233, 8233}
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D4"), RowCol.Rows);

        // Get first series
        ISeries series1 = shape.getChart().getSeriesCollection().get(0);

        // Set HasErrorBars as true
        series1.setHasErrorBars(true);

        // Set value type of first series' error bar as custom
        series1.getYErrorBar().setValueType(ErrorBarType.Custom);
        series1.getYErrorBar().setPlus("={2000}");
        series1.getYErrorBar().setMinus("={1000}");
    }
}
