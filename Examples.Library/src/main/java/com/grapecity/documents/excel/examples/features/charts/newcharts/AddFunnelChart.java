package com.grapecity.documents.excel.examples.features.charts.newcharts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddFunnelChart extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:B7").setValue(new Object[][]{
        	{"Stage", "Amount"},
            {"Prospects", 500},
            {"Qualified prospects", 425},
            {"Needs analysis", 200},
            {"Price quotes", 150},
            {"Negotiations", 100},
            {"Closed sales", 90}
        });
        worksheet.getRange("A:A").getColumns().autoFit();

        //Create a funnel chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.Funnel, 300, 20, 300, 200);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:B7"));

        //Set the axis invisible.
        IAxis axis = shape.getChart().getAxes().item(AxisType.Category, AxisGroup.Primary);
        axis.setVisible(false);
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
