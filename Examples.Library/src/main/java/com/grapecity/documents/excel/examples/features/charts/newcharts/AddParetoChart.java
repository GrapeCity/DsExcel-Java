package com.grapecity.documents.excel.examples.features.charts.newcharts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.BinsType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddParetoChart extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:B11").setValue(new Object[][]{
        	  {"Complaint", "Count"},
              {"Too noisy", 27},
              {"Overpriced", 789},
              {"Food is tasteless", 65},
              {"Food is not fresh", 9},
              {"Food is too salty", 15},
              {"Not clean", 30},
              {"Unfriendly staff", 12},
              {"Wait time", 109},
              { "No atmosphere", 45},
              {"Small portions", 621 }
        });
        worksheet.getRange("A:A").getColumns().autoFit();

        //Create a pareto chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.Pareto, 300, 20, 300, 200);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B11"));

        //Sets bins type by size.
        shape.getChart().getChartGroups().get(0).setBinsType(BinsType.BinsTypeBinSize);
        shape.getChart().getChartGroups().get(0).setBinWidthValue(300);

        //Set underflow bin value.
        shape.getChart().getChartGroups().get(0).setBinsUnderflowEnabled(true);
        shape.getChart().getChartGroups().get(0).setBinsOverflowValue(50);
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
