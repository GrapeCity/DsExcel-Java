package com.grapecity.documents.excel.examples.exceltemplates;

import java.net.URL;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ITextRange;
import com.grapecity.documents.excel.drawing.SolidColorType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BloodSugarTracker extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Load template file Blood sugar tracker.xlsx from resource
        workbook.open(this.getResourceStream("xlsx/Blood sugar tracker.xlsx"));

        IWorksheet worksheet = workbook.getActiveSheet();

        //insert 19 rows  
        worksheet.getRange("1:19").insert();

        //Change the rows(2~5) RowHeight
        worksheet.getRows().get(1).setRowHeight(34.5);
        worksheet.getRows().get(2).setRowHeight(15.75);
        worksheet.getRows().get(3).setRowHeight(19.5);
        worksheet.getRows().get(4).setRowHeight(15.75);

        //Set values
        worksheet.getRange("B2").setValue("BLOOD SUGAR TRACKING");
        worksheet.getRange("B4").setValue("CHARTED PROGRESS");

        //Set Styles
        worksheet.getRange("B2").getFont().setThemeFont(ThemeFont.Major);
        worksheet.getRange("B2").getFont().setSize(26);
        worksheet.getRange("B2").getFont().setThemeColor(ThemeColor.Dark1);
        worksheet.getRange("B2").getFont().setTintAndShade(0.34998626667073579);
        worksheet.getRange("B2:D2").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        worksheet.getRange("B2:D2").getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent1);

        worksheet.getRange("B4").getFont().setThemeFont(ThemeFont.Major);
        worksheet.getRange("B4").getFont().setBold(true);
        worksheet.getRange("B4").getFont().setSize(14);
        worksheet.getRange("B4").getFont().setThemeColor(ThemeColor.Dark1);
        worksheet.getRange("B4").getFont().setTintAndShade(0.34998626667073579);
        worksheet.getRange("B4:D4").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        worksheet.getRange("B4:D4").getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent2);

        //Add chart
        IShape shape = worksheet.getShapes().addChart(ChartType.Line, 9.75, 100.5, 365, 203.25);
        shape.setName("BloodSugarProgress");

        //Add Series.
        ISeries series1 = shape.getChart().getSeriesCollection().newSeries();
        series1.setFormula("=SERIES('BLOOD SUGAR DATA'!$C$23,'BLOOD SUGAR DATA'!$B$24:$B$45,'BLOOD SUGAR DATA'!$C$24:$C$45,1)");
        series1.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent1);
        series1.getFormat().getLine().setWeight(2.5);

        ISeries series2 = shape.getChart().getSeriesCollection().newSeries();
        series2.setFormula("=SERIES('BLOOD SUGAR DATA'!$D$23,'BLOOD SUGAR DATA'!$B$24:$B$45,'BLOOD SUGAR DATA'!$D$24:$D$45,2)");
        series2.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent2);
        series2.getFormat().getLine().setWeight(2.5);

        //Hidden the chart title
        shape.getChart().setHasTitle(false);

        //Hidden the category axis
        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);
        category_axis.setVisible(false);

        //Set value axis units
        IAxis value_axis = shape.getChart().getAxes().item(AxisType.Value);
        value_axis.setMaximumScale(140);
        value_axis.setMinimumScale(50);
        value_axis.setMajorUnit(10);
        value_axis.setMinorUnit(2);

        //Add a rectange shape
        IShape shape1 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 385, 84.75, 102, 218.25);
        shape1.getFill().solid();
        shape1.getFill().getColor().setObjectThemeColor(ThemeColor.Accent1);
        shape1.getFill().getColor().setBrightness(0.6);

        //set shape's border to no line
        shape1.getLine().getColor().setColorType(SolidColorType.None);

        //set shape rich text
        ITextRange shape1_p1 = shape1.getTextFrame().getTextRange().getParagraphs().get(0);
        shape1_p1.setText("INFO:");
        shape1_p1.getRuns().add(" Blood sugar levels will vary from person-to-person.  There are many factors to keeping it within your normal range and isn't based on sugar alone.  Consult a physician for additional information or follow-up.");

        ITextRange shape1_p2 = shape1.getTextFrame().getTextRange().getParagraphs().add("");
        ITextRange shape1_p3 = shape1.getTextFrame().getTextRange().getParagraphs().add("More info can be found here:");
        ITextRange shape1_p4 = shape1.getTextFrame().getTextRange().getParagraphs().add("http://diabetes.webmd.com/blood-glucose");

        shape1.getTextFrame().getTextRange().getFont().setSize(10);
        shape1.getTextFrame().getTextRange().getFont().setThemeFont(ThemeFont.Minor);
        shape1.getTextFrame().getTextRange().getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape1.getTextFrame().getTextRange().getFont().getColor().setBrightness(0.25);
        shape1_p1.getRuns().get(0).getFont().setBold(true);
        shape1_p3.getRuns().get(0).getFont().setBold(true);

        //Do table filter
        ITable table = worksheet.getTables().get(0);
        table.getRange().autoFilter(1, ">=102");

    }

    @Override
    public String getTemplateName() {

        return "Blood sugar tracker.xlsx";

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public String[] getResources() {
        return new String[]{"xlsx/Blood sugar tracker.xlsx"};
    }
}
