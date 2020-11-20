package com.grapecity.documents.excel.examples.features.charts.chartgallery;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.CategoryType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.MarkerStyle;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.TickMark;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Stock_HighLowCloseStockChart extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addChart(ChartType.StockHLC, 350, 20, 360, 230);
        worksheet.getRange("A1:D17").setValue(new Object[][]{
                {null, "High", "Low", "Close"},
                {new GregorianCalendar(2019, 8, 1), 105.76, 92.38, 100.94},
                {new GregorianCalendar(2019, 8, 2), 102.45, 90.14, 93.45},
                {new GregorianCalendar(2019, 8, 3), 102.11, 85.01, 99.89},
                {new GregorianCalendar(2019, 8, 4), 106.01, 94.04, 99.45},
                {new GregorianCalendar(2019, 8, 5), 108.23, 98.16, 104.33},
                {new GregorianCalendar(2019, 8, 8), 107.7, 91.02, 102.17},
                {new GregorianCalendar(2019, 8, 9), 110.36, 101.62, 110.07},
                {new GregorianCalendar(2019, 8, 10), 115.97, 106.89, 112.39},
                {new GregorianCalendar(2019, 8, 11), 120.32, 112.15, 117.52},
                {new GregorianCalendar(2019, 8, 12), 122.03, 114.67, 114.75},
                {new GregorianCalendar(2019, 8, 15), 120.46, 106.21, 116.85},
                {new GregorianCalendar(2019, 8, 16), 118.08, 113.55, 116.69},
                {new GregorianCalendar(2019, 8, 17), 128.23, 110.91, 117.25},
                {new GregorianCalendar(2019, 8, 18), 120.55, 108.09, 112.52},
                {new GregorianCalendar(2019, 8, 19), 112.58, 105.42, 109.12},
                {new GregorianCalendar(2019, 8, 22), 115.23, 97.25, 101.56},
        });
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D17"), RowCol.Columns);
        shape.getChart().getChartTitle().setText("High-Low-Close Stock Chart");
        IAxis valueAxis = shape.getChart().getAxes().item(AxisType.Value);
        IAxis categoryAxis = shape.getChart().getAxes().item(AxisType.Category);
        ISeries series_close = shape.getChart().getSeriesCollection().get(2);
        //config value axis
        valueAxis.setMinimumScale(80);
        valueAxis.setMaximumScale(140);
        valueAxis.setMajorUnit(15);
        //config category axis
        categoryAxis.setCategoryType(CategoryType.CategoryScale);
        categoryAxis.setMajorTickMark(TickMark.Outside);
        categoryAxis.setTickLabelSpacingIsAuto(false);
        categoryAxis.setTickLabelSpacing(5);
        series_close.setMarkerStyle(MarkerStyle.Square);

    }

}
