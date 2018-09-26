package com.grapecity.documents.excel.examples.excelreporting;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.FormatConditionOperator;
import com.grapecity.documents.excel.FormatConditionType;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IFormatCondition;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.TotalsCalculation;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ArrowheadLength;
import com.grapecity.documents.excel.drawing.ArrowheadStyle;
import com.grapecity.documents.excel.drawing.ArrowheadWidth;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.AxisGroup;
import com.grapecity.documents.excel.drawing.AxisType;
import com.grapecity.documents.excel.drawing.CategoryType;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.GradientStyle;
import com.grapecity.documents.excel.drawing.IAxis;
import com.grapecity.documents.excel.drawing.ISeries;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ITextRange;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.LineStyle;
import com.grapecity.documents.excel.drawing.SolidColorType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BloodPressureTracker extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //theme
        //create a custom theme.
        ITheme theme = new Theme("testTheme");
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.FromArgb(222, 222, 212));
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.FromArgb(30, 46, 47));
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.FromArgb(233, 117, 90));
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.FromArgb(122, 182, 186));
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.FromArgb(125, 181, 135));
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.FromArgb(230, 191, 94));
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.FromArgb(230, 143, 77));
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.FromArgb(194, 107, 112));
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.FromArgb(122, 182, 186));
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.FromArgb(166, 140, 177));
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Gill Sans");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Gill Sans");

        //assign the custom theme for workbook.
        workbook.setTheme(theme);

        //does not show sheet gridlines.
        worksheet.getSheetView().setDisplayGridlines(false);

        //RowHeightColumnWidth
        //set row height and column width.
        worksheet.setStandardHeight(12.75);
        worksheet.setStandardWidth(8.43);
        worksheet.getRows().get(1).setRowHeight(32.25);
        worksheet.getRows().get(2).setRowHeight(13.5);
        worksheet.getRows().get(3).setRowHeight(18.75);
        worksheet.getRows().get(6).setRowHeight(18.75);
        worksheet.getRows().get(9).setRowHeight(18.75);
        worksheet.getRows().get(12).setRowHeight(18.75);
        worksheet.getRows().get(15).setRowHeight(19.5);
        worksheet.getRows().get(16).setRowHeight(13.5);
        worksheet.getRows().get(33).setRowHeight(19.5);
        worksheet.getRows().get(34).setRowHeight(13.5);

        worksheet.getColumns().get(0).setColumnWidth(1.7109375);
        worksheet.getColumns().get(1).setColumnWidth(12.140625);
        worksheet.getColumns().get(2).setColumnWidth(12.140625);
        worksheet.getColumns().get(3).setColumnWidth(12.140625);
        worksheet.getColumns().get(4).setColumnWidth(11.85546875);
        worksheet.getColumns().get(5).setColumnWidth(12.7109375);
        worksheet.getColumns().get(6).setColumnWidth(13.85546875);
        worksheet.getColumns().get(7).setColumnWidth(44.7109375);

        //Values
        //initialize worksheet's values.
        worksheet.setName("BLOOD PRESSURE DATA");
        worksheet.getRange("B2").setValue("BLOOD PRESSURE TRACKER");
        worksheet.getRange("B4:F13").setValue(new Object[][]{
            {"NAME", null, null, null, null},
            {null, null, null, null, null},
            {null, null, null, "Systolic", "Diastolic"},
            {"TARGET BLOOD PRESSURE", null, null, 120, 80},
            {null, null, null, null, null},
            {null, null, null, "Systolic", "Diastolic"},
            {"CALL PHYSICIAN IF ABOVE", null, null, 140, 90},
            {null, null, null, null, null},
            {null, null, null, null, null},
            {"PHYSICIAN PHONE NUMBER", null, null, "[Phone Number]", null}
        });
        worksheet.getRange("B16").setValue("CHARTED PROGRESS");
        worksheet.getRange("B34").setValue("DATA ENTRY");

        //Table
        //initialize table data.
        worksheet.getRange("B36:H44").setValue(new Object[][]{
            {"TIME", "DATE", "AM/PM", "SYSTOLIC", "DIASTOLIC", "HEART RATE", "NOTES"},
            {new GregorianCalendar(1899, 11, 30, 10, 00, 00), new GregorianCalendar(2013, 6, 1), "AM", 129, 99, 72, null},
            {new GregorianCalendar(1899, 11, 30, 18, 00, 00), new GregorianCalendar(2013, 6, 1), "PM", 133, 80, 75, null},
            {new GregorianCalendar(1899, 11, 30, 10, 30, 00), new GregorianCalendar(2013, 6, 2), "AM", 142, 86, 70, null},
            {new GregorianCalendar(1899, 11, 30, 19, 00, 00), new GregorianCalendar(2013, 6, 2), "PM", 141, 84, 68, null},
            {new GregorianCalendar(1899, 11, 30, 9,  00, 00), new GregorianCalendar(2013, 6, 2), "AM", 137, 84, 70, null},
            {new GregorianCalendar(1899, 11, 30, 18, 30, 00), new GregorianCalendar(2013, 6, 3), "PM", 139, 83, 72, null},
            {new GregorianCalendar(1899, 11, 30, 10, 00, 00), new GregorianCalendar(2013, 6, 4), "AM", 140, 85, 78, null},
            {new GregorianCalendar(1899, 11, 30, 18, 00, 00), new GregorianCalendar(2013, 6, 4), "PM", 138, 85, 69, null},
        });
        ITable table = worksheet.getTables().add(worksheet.getRange("B36:H44"), true);
        table.setShowTotals(true);

        //set total row formulas.
        table.getColumns().get(0).getTotal().setValue("Average");
        table.getColumns().get(3).setTotalsCalculation(TotalsCalculation.Average);
        table.getColumns().get(4).setTotalsCalculation(TotalsCalculation.Average);
        table.getColumns().get(5).setTotalsCalculation(TotalsCalculation.Average);
        table.getColumns().get(6).setTotalsCalculation(TotalsCalculation.None);

        //config data body range and total range's number format.
        table.getColumns().get(0).getDataBodyRange().setNumberFormat("h:mm;@");
        table.getColumns().get(1).getDataBodyRange().setNumberFormat("m/d/yyyy");
        table.getColumns().get(3).getDataBodyRange().setNumberFormat("0");
        table.getColumns().get(4).getDataBodyRange().setNumberFormat("0");
        table.getColumns().get(5).getDataBodyRange().setNumberFormat("0");
        table.getColumns().get(3).getTotal().setNumberFormat("0");
        table.getColumns().get(4).getTotal().setNumberFormat("0");
        table.getColumns().get(5).getTotal().setNumberFormat("0");

        //config table range's alignment.
        table.getRange().setHorizontalAlignment(HorizontalAlignment.Left);
        table.getRange().setIndentLevel(0);
        table.getRange().setVerticalAlignment(VerticalAlignment.Center);

        //TableStyle
        //create a custom table style.
        ITableStyle tablestyle = workbook.getTableStyles().add("testStyle");
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setThemeColor(ThemeColor.Dark1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setTintAndShade(0.25);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thin);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeTop).setTintAndShade(0.4);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setTintAndShade(0.4);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setTintAndShade(0.4);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setTintAndShade(0.4);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setTintAndShade(0.4);

        tablestyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe).getInterior().setThemeColor(ThemeColor.Accent1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe).getInterior().setTintAndShade(0.8);

        tablestyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getFont().setBold(true);
        tablestyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getFont().setThemeColor(ThemeColor.Dark1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getFont().setTintAndShade(0.25);
        tablestyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setThemeColor(ThemeColor.Accent1);

        tablestyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getFont().setBold(true);
        tablestyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getFont().setThemeColor(ThemeColor.Dark1);
        tablestyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getFont().setTintAndShade(0.25);
        tablestyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Double);
        tablestyle.getTableStyleElements().get(TableStyleElementType.TotalRow).getBorders().get(BordersIndex.EdgeTop).setThemeColor(ThemeColor.Accent1);

        //assign custom table style for table.
        table.setTableStyle(workbook.getTableStyles().get("testStyle"));

        //Style
        //assign built-in styles for ranges.
        worksheet.getRange("B2:H2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("B4:F4, B7:D7, B10:D10, B13:D13").setStyle(workbook.getStyles().get("Heading 2"));
        worksheet.getRange("B16:H16, B34:H34").setStyle(workbook.getStyles().get("Heading 3"));

        //modify built-in styles.
        IStyle style_Heading1 = workbook.getStyles().get("Heading 1");
        style_Heading1.setHorizontalAlignment(HorizontalAlignment.General);
        style_Heading1.setVerticalAlignment(VerticalAlignment.Center);
        style_Heading1.getFont().setThemeFont(ThemeFont.Major);
        style_Heading1.getFont().setSize(24);
        style_Heading1.getFont().setBold(true);
        style_Heading1.getFont().setThemeColor(ThemeColor.Accent1);
        style_Heading1.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        style_Heading1.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent1);

        style_Heading1.setIncludeAlignment(true);
        style_Heading1.setIncludeFont(true);
        style_Heading1.setIncludeBorder(true);
        style_Heading1.setIncludeNumber(false);
        style_Heading1.setIncludePatterns(false);
        style_Heading1.setIncludeProtection(false);

        IStyle style_Heading2 = workbook.getStyles().get("Heading 2");
        style_Heading2.setHorizontalAlignment(HorizontalAlignment.General);
        style_Heading2.setVerticalAlignment(VerticalAlignment.Bottom);
        style_Heading2.getFont().setThemeFont(ThemeFont.Minor);
        style_Heading2.getFont().setSize(14);
        style_Heading2.getFont().setThemeColor(ThemeColor.Dark1);
        style_Heading2.getFont().setTintAndShade(0.25);
        style_Heading2.getFont().setBold(false);
        style_Heading2.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Dotted);
        style_Heading2.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Light1);
        style_Heading2.getBorders().get(BordersIndex.EdgeBottom).setTintAndShade(-0.5);

        style_Heading2.setIncludeAlignment(true);
        style_Heading2.setIncludeFont(true);
        style_Heading2.setIncludeBorder(true);
        style_Heading2.setIncludeNumber(false);
        style_Heading2.setIncludePatterns(false);
        style_Heading2.setIncludeProtection(false);

        IStyle style_Heading3 = workbook.getStyles().get("Heading 3");
        style_Heading3.setHorizontalAlignment(HorizontalAlignment.General);
        style_Heading3.setVerticalAlignment(VerticalAlignment.Center);
        style_Heading3.getFont().setThemeFont(ThemeFont.Minor);
        style_Heading3.getFont().setSize(14);
        style_Heading3.getFont().setBold(true);
        style_Heading3.getFont().setThemeColor(ThemeColor.Dark1);
        style_Heading3.getFont().setTintAndShade(0.25);
        style_Heading3.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        style_Heading3.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent2);

        style_Heading3.setIncludeAlignment(true);
        style_Heading3.setIncludeFont(true);
        style_Heading3.setIncludeBorder(true);
        style_Heading3.setIncludeNumber(false);
        style_Heading3.setIncludePatterns(false);
        style_Heading3.setIncludeProtection(false);

        IStyle style_Normal = workbook.getStyles().get("Normal");
        style_Normal.setNumberFormat("General");
        style_Normal.setHorizontalAlignment(HorizontalAlignment.General);
        style_Normal.setVerticalAlignment(VerticalAlignment.Center);
        style_Normal.getFont().setThemeFont(ThemeFont.Minor);
        style_Normal.getFont().setSize(10);
        style_Normal.getFont().setThemeColor(ThemeColor.Dark1);
        style_Normal.getFont().setTintAndShade(0.25);

        style_Normal.setIncludeAlignment(true);
        style_Normal.setIncludeFont(true);
        style_Normal.setIncludeBorder(true);
        style_Normal.setIncludeNumber(true);
        style_Normal.setIncludePatterns(true);
        style_Normal.setIncludeProtection(true);

        //modify cell styles.
        worksheet.getRange("B4").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        worksheet.getRange("C4:F4").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        IRange range1 = worksheet.getRange("E7:F7, E10:F10");
        range1.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Dotted);
        range1.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Light1);
        range1.getBorders().get(BordersIndex.EdgeBottom).setTintAndShade(-0.5);
        range1.getFont().setBold(true);
        IRange range2 = worksheet.getRange("E7, E10");
        range2.getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        range2.getBorders().get(BordersIndex.EdgeRight).setThemeColor(ThemeColor.Light1);
        range2.getBorders().get(BordersIndex.EdgeRight).setTintAndShade(-0.5);

        //Chart
        //create a new chart.
        IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 8.99984251968504, 268.5, 627.750157480315, 184.5);

        //create series for chart.
        ISeries series_systolic = shape.getChart().getSeriesCollection().newSeries();
        ISeries series_diatolic = shape.getChart().getSeriesCollection().newSeries();
        ISeries series_HeartRate = shape.getChart().getSeriesCollection().newSeries();

        //set series formulas.
        series_systolic.setFormula("=SERIES('BLOOD PRESSURE DATA'!$E$36,'BLOOD PRESSURE DATA'!$C$37:$D$44,'BLOOD PRESSURE DATA'!$E$37:$E$44,1)");
        series_diatolic.setFormula("=SERIES('BLOOD PRESSURE DATA'!$F$36,'BLOOD PRESSURE DATA'!$C$37:$D$44,'BLOOD PRESSURE DATA'!$F$37:$F$44,2)");

        //set series plot on secondary axis, and change its chart type.
        series_HeartRate.setAxisGroup(AxisGroup.Secondary);
        series_HeartRate.setChartType(ChartType.Line);
        series_HeartRate.setFormula("=SERIES('BLOOD PRESSURE DATA'!$G$36,,'BLOOD PRESSURE DATA'!$G$37:$G$44,3)");

        //set series fill to gradient fill.
        series_systolic.getFormat().getFill().twoColorGradient(GradientStyle.Horizontal, 1);
        series_systolic.getFormat().getFill().setGradientAngle(270);
        series_systolic.getFormat().getFill().getGradientStops().get(0).getColor().setRGB(Color.FromArgb(255, 172, 175));
        series_systolic.getFormat().getFill().getGradientStops().get(1).getColor().setRGB(Color.FromArgb(255, 227, 228));
        series_systolic.getFormat().getFill().getGradientStops().insert(0xFEC6C8, 0.35);
        series_systolic.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent6);

        series_diatolic.getFormat().getFill().twoColorGradient(GradientStyle.Horizontal, 1);
        series_diatolic.getFormat().getFill().setGradientAngle(270);
        series_diatolic.getFormat().getFill().getGradientStops().get(0).getColor().setRGB(Color.FromArgb(255, 192, 147));
        series_diatolic.getFormat().getFill().getGradientStops().get(1).getColor().setRGB(Color.FromArgb(255, 227, 212));
        series_diatolic.getFormat().getFill().getGradientStops().insert(0xFFCBA9, 0.35);
        series_diatolic.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent5);

        //set series gap width and overlap.
        shape.getChart().getColumnGroups().get(0).setGapWidth(150);
        shape.getChart().getColumnGroups().get(0).setOverlap(0);

        //set series line style.
        series_HeartRate.getFormat().getLine().setBeginArrowheadLength(ArrowheadLength.Medium);
        series_HeartRate.getFormat().getLine().setBeginArrowheadStyle(ArrowheadStyle.None);
        series_HeartRate.getFormat().getLine().setBeginArrowheadWidth(ArrowheadWidth.Medium);
        series_HeartRate.getFormat().getLine().getColor().setObjectThemeColor(ThemeColor.Accent4);
        series_HeartRate.getFormat().getLine().getColor().setTintAndShade(0);
        series_HeartRate.getFormat().getLine().setDashStyle(LineDashStyle.Solid);
        series_HeartRate.getFormat().getLine().setEndArrowheadLength(ArrowheadLength.Medium);
        series_HeartRate.getFormat().getLine().setEndArrowheadStyle(ArrowheadStyle.None);
        series_HeartRate.getFormat().getLine().setEndArrowheadWidth(ArrowheadWidth.Medium);
        series_HeartRate.getFormat().getLine().setStyle(LineStyle.Single);
        series_HeartRate.getFormat().getLine().setWeight(1.25);

        IAxis primary_axis = shape.getChart().getAxes().item(AxisType.Value, AxisGroup.Primary);
        primary_axis.setHasTitle(true);
        primary_axis.getAxisTitle().setText("BLOOD PRESSURE");
        primary_axis.getAxisTitle().setIncludeInLayout(true);

        IAxis secondary_axis = shape.getChart().getAxes().item(AxisType.Value, AxisGroup.Secondary);
        secondary_axis.setHasTitle(true);
        secondary_axis.getAxisTitle().setText("HEART RATE");
        secondary_axis.getAxisTitle().setIncludeInLayout(true);

        IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category, AxisGroup.Primary);
        category_axis.setHasTitle(true);
        category_axis.setCategoryType(CategoryType.CategoryScale);
        category_axis.getFormat().getLine().getColor().setColorType(SolidColorType.None);

        shape.getChart().setHasTitle(false);
        //set chart font style.
        shape.getChart().getChartArea().getFont().setSize(9);
        shape.getChart().getChartArea().getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape.getChart().getChartArea().getFont().getColor().setBrightness(0.5);


        //Shape
        IShape shape1 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 402, 77.25, 234, 100);
        shape1.getFill().solid();
        shape1.getFill().getColor().setObjectThemeColor(ThemeColor.Accent1);
        shape1.getFill().getColor().setBrightness(0.6);
        //set shape's border to no line.
        shape1.getLine().getColor().setColorType(SolidColorType.None);

        //set shape rich text.
        ITextRange shape1_p1 = shape1.getTextFrame().getTextRange().getParagraphs().get(0);
        shape1_p1.setText("*");
        shape1_p1.getRuns().add(" Blood pressures may vary dependent on many");
        shape1_p1.getRuns().add(" factors.  Always consult with a physician about what is normal for you.  These numbers may vary slightly.");

        ITextRange shape1_p2 = shape1.getTextFrame().getTextRange().getParagraphs().add("");
        ITextRange shape1_p3 = shape1.getTextFrame().getTextRange().getParagraphs().add("Info from National Institute of Health:");
        ITextRange shape1_p4 = shape1.getTextFrame().getTextRange().getParagraphs().add("http://www.nhlbi.nih.gov/health/health-topics/topics/hbp/");

        shape1.getTextFrame().getTextRange().getFont().setSize(10);
        shape1.getTextFrame().getTextRange().getFont().setThemeFont(ThemeFont.Minor);
        shape1.getTextFrame().getTextRange().getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape1.getTextFrame().getTextRange().getFont().getColor().setBrightness(0.25);
        shape1_p3.getRuns().get(0).getFont().setBold(true);

        IShape shape2 = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 421.5, 546.75, 198, 50);
        shape2.getFill().solid();
        shape2.getFill().getColor().setObjectThemeColor(ThemeColor.Accent3);
        shape2.getFill().getColor().setBrightness(0.6);
        //set shape's border to no line.
        shape2.getLine().getColor().setColorType(SolidColorType.None);

        ITextRange shape2_p1 = shape2.getTextFrame().getTextRange().getParagraphs().get(0);
        shape2_p1.setText("NOTE:");
        shape2_p1.getRuns().add(" Any blood pressure readings over the indicated numbers (\"CALL PHYSICIAN IF ABOVE\") will be");
        shape2_p1.getRuns().add(" highlighted.");

        shape2.getTextFrame().getTextRange().getFont().setSize(10);
        shape2.getTextFrame().getTextRange().getFont().setThemeFont(ThemeFont.Minor);
        shape2.getTextFrame().getTextRange().getFont().getColor().setObjectThemeColor(ThemeColor.Dark1);
        shape2.getTextFrame().getTextRange().getFont().getColor().setBrightness(0.25);
        shape2_p1.getRuns().get(0).getFont().setBold(true);

        //DefinedName
        //create defined names for workbook.
        workbook.getNames().add("MaxDiastolic", "='BLOOD PRESSURE DATA'!$F$10");
        workbook.getNames().add("MaxSystolic", "='BLOOD PRESSURE DATA'!$E$10");

        //ConditionalFormat
        //create conditional format for ranges.
        IFormatCondition condition1 = (IFormatCondition) worksheet.getRange("E37:E44").getFormatConditions().add(FormatConditionType.Expression, FormatConditionOperator.Between, "=E37>MaxSystolic", null);
        IFormatCondition condition2 = (IFormatCondition) worksheet.getRange("F37:F44").getFormatConditions().add(FormatConditionType.Expression, FormatConditionOperator.Between, "=F37>MaxDiastolic", null);
        condition1.getInterior().setColor(Color.GetRed());
        condition2.getInterior().setColor(Color.GetRed());

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }
}
