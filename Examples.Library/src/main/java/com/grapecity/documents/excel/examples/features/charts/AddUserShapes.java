package com.grapecity.documents.excel.examples.features.charts;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddUserShapes extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        worksheet.getRange("A1:C10").setValue(new Object[][]{
            {"Task", "Worker 1", "Worker 2"},
            {"Task 1", 7, 10},
            {"Task 2", 5, 1},
            {"Task 3", 3, 6},
            {"Task 4", 10, 5},
            {"Task 5", 4, 4},
            {"Task 6", 5, 8},
            {"Task 7", 8, 7},
            {"Task 8", 2, 5},
            {"Task 9", 6, 4}
        });

        //Create a funnel chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 250, 20, 400, 250);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:C10"));
        shape.getChart().setHasLegend(false);
        shape.getChart().getChartTitle().setText(" ");
        
        //Add a bussiness logo in the line chart area
        InputStream stream = this.getResourceStream("logo.png");
        try {
        	shape.getChart().addPicture(stream, ImageType.PNG, 170, 10, 60, 10);
        } catch (IOException e) {
        	e.printStackTrace();
        }

        //Add shapes in the line chart area 
        IShape userShape1 = shape.getChart().addShape(AutoShapeType.Rectangle, 30, 45, 60, 20);
        userShape1.getFill().getColor().setObjectThemeColor(ThemeColor.Accent2);
        userShape1.getLine().getColor().setObjectThemeColor(ThemeColor.Accent2);
        userShape1.getTextFrame().getTextRange().get(0).setText("Worker 2");

        IShape userShape2 = shape.getChart().addShape(AutoShapeType.Rectangle, 330, 110, 60, 20);
        userShape2.getFill().getColor().setObjectThemeColor(ThemeColor.Accent1);
        userShape2.getLine().getColor().setObjectThemeColor(ThemeColor.Accent1);
        userShape2.getTextFrame().getTextRange().get(0).setText("Worker 1");
	}

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
    }
	
	@Override
	public boolean getShowScreenshot() {
        return true;
	}
}
