package com.grapecity.documents.excel.examples.features.charts.newcharts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ParentDataLabelOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddTreeMapChart extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
		
        worksheet.getRange("A1:D16").setValue(new Object[][]{
        	 {"Quarter", "Month", "Week", "Output"},
             {"1st", "Jan", null, 3.5},
             {null, "Feb", "Week1", 1.2},
             {null, null, "Week2", 0.8},
             {null, null, "Week3", 0.6},
             {null, null, "Week4", 0.5},
             {null, "Mar", null, 1.7 },
             {"2st", "Apr", null, 1.1},
             {null, "May", null, 0.8},
             {null, "Jun", null, 0.3},
             {"3st", "July", null, 0.7},
             {null, "Aug", null, 0.6},
             {null, "Sept", null, 0.1},
             {"4st", "Oct", null, 0.5},
             {null, "Nov", null, 0.4},
             {null, "Dec", null, 0.3},
        });

        //Create a treemap chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.Treemap, 300, 20, 300, 200);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D16"));
        
        //Set the parent data labels are displayed as banners.
        shape.getChart().getSeriesCollection().get(0).setParentDataLabelOption(ParentDataLabelOptions.Banner);

        //Modify chart title text.
        shape.getChart().getChartTitle().setText("Annual Report");
	}

	@Override
	public boolean getIsNew() {
		return true;
	}
}
