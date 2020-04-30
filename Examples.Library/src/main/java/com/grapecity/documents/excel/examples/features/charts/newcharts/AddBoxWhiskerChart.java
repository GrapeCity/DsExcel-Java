package com.grapecity.documents.excel.examples.features.charts.newcharts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddBoxWhiskerChart extends ExampleBase {
   
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:D16").setValue(new Object[][]{
            {"Course", "SchoolA", "SchoolB", "SchoolC"},
            {"English", 63, 53, 45},
            {"Physics", 61, 55, 65},
            {"English", 63, 50, 65},
            {"Math", 62, 51, 64},
            {"English", 46, 53, 66},
            {"English", 58, 56, 67},
            {"Math", 60, 51, 67},
            {"Math", 62, 53, 66},
            {"English", 63, 54, 64},
            {"English", 63, 52, 67},
            {"Physics", 60, 56, 64},
            {"English", 60, 56, 67},
            {"Math", 61, 56, 45},
            {"Math", 63, 58, 64},
            {"English", 59, 54, 65}
        });

        //Create a box&whisker chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.BoxWhisker, 300, 20, 300, 200);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D16"));

        //Config value axis's scale.
        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value, AxisGroup.Primary);
        value_axis.setMinimumScale(40);
        value_axis.setMaximumScale(70);

        //Config the display of box&whisker plot.  
        ISeries series = shape.getChart().getSeriesCollection().get(0);
        series.setShowInnerPoints(true);
        series.setShowOutlierPoints(false);
        series.setShowMeanMarkers(false);
        series.setShowMeanLine(true);
        series.setQuartileCalculationInclusiveMedian(true);
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
