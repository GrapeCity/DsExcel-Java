package com.grapecity.documents.excel.examples.features.charts.axes;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.*;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class ConfigTickMarkLabelAngle extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Manufacturing output chart.xlsx");
        workbook.open(fileStream);

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShape shape = worksheet.getShapes().get(0);

        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);

        //config tick label's angle
        category_axis.getTickLabels().setOrientation(-45);
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Manufacturing output chart.xlsx" };
    }
}
