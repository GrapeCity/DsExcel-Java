package com.grapecity.documents.excel.examples.features.charts.newcharts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.ILegendEntries;
import com.grapecity.documents.excel.drawing.IPoints;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddWaterfallChart extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:B8").setValue(new Object[][]{
        	 {"Starting Amt", 130},
             {"Measurement 1", 25},
             {"Measurement 2", -75},
             {"Subtotal", 80},
             {"Measurement 3", 45},
             {"Measurement 4", -65},
             {"Measurement 5", 80},
             {"Total", 140}
        });
        worksheet.getRange("A:A").getColumns().autoFit();

        //Create a waterfall chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.Waterfall, 300, 20, 300, 250);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B8"));

        //Set subtotal&total points.
        IPoints points = shape.getChart().getSeriesCollection().get(0).getPoints();
        points.get(3).setIsTotal(true);
        points.get(7).setIsTotal(true);

        //Connector lines are not shown.
        ISeries series = shape.getChart().getSeriesCollection().get(0);
        series.setShowConnectorLines(false);
        
        //Modify the fill color of the first legend entry.
        ILegendEntries legendEntries = shape.getChart().getLegend().getLegendEntries();
        legendEntries.get(0).getFormat().getFill().getColor().setObjectThemeColor(ThemeColor.Accent6);
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}
	
	@Override
	public boolean getShowScreenshot() {
        return true;
	}
}
