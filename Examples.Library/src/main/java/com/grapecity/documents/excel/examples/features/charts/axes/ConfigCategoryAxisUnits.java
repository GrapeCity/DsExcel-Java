package com.grapecity.documents.excel.examples.features.charts.axes;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

import com.grapecity.documents.excel.DateInfo;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.drawing.TimeUnit;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigCategoryAxisUnits extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A2:A6").setNumberFormat("m/d/yyyy");
        worksheet.getRange("A1:D6").setValue(new Object[][]
                {
                        {null, "S1", "S2", "S3"},
                        {new GregorianCalendar(2015, 9, 7), 10, 25, 25},
                        {new GregorianCalendar(2015, 9, 24), 51, 36, 27},
                        {new GregorianCalendar(2015, 10, 8), 52, 85, 30},
                        {new GregorianCalendar(2015, 10, 25), 22, 65, 65},
                        {new GregorianCalendar(2015, 11, 10), 23, 69, 69}
                });

        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 250, 20, 360, 230);
        shape.getChart().getSeriesCollection().add(worksheet.getRange("A1:D6"), RowCol.Columns, true, true);

        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);
        // DateTime(2015, 12, 20)
        category_axis.setMaximumScale(DateInfo.ToOADate(new GregorianCalendar(2015, 11, 20)));
        // DateTime(2015, 10, 1)
        category_axis.setMinimumScale(DateInfo.ToOADate(new GregorianCalendar(2015, 9, 1)));
        category_axis.setBaseUnit(TimeUnit.Months);
        category_axis.setMajorUnitScale(TimeUnit.Months);
        category_axis.setMajorUnit(1);
        category_axis.setMinorUnitScale(TimeUnit.Days);
        category_axis.setMinorUnit(15);
    }

}
