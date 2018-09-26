package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ConditionValueTypes;
import com.grapecity.documents.excel.DataBarFillType;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.FormatConditionOperator;
import com.grapecity.documents.excel.FormatConditionType;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IDataBar;
import com.grapecity.documents.excel.IFormatCondition;
import com.grapecity.documents.excel.IIconSetCondition;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IValidation;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.IconSetType;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.ValidationAlertStyle;
import com.grapecity.documents.excel.ValidationOperator;
import com.grapecity.documents.excel.ValidationType;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ToDoListWithProgressTracker extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object[] data = new Object[][]{
                {"TASK", "PRIORITY", "STATUS", "START DATE", "DUE DATE", "% COMPLETE", "DONE?", "NOTES"},
                {"First Thing I Need To Do", "Normal", "Not Started", null, null, 0, null, null},
                {"Other Thing I Need To Finish", "High", "In Progress", null, null, 0.5, null, null},
                {"Something Else To Get Done", "Low", "Complete", null, null, 1, null, null},
                {"More Errands And Things", "Normal", "In Progress", null, null, 0.75, null, null},
                {"So Much To Get Done This Week", "High", "In Progress", null, null, 0.25, null, null}
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("To-Do List");
        worksheet.setTabColor(Color.FromArgb(148, 112, 135));
        worksheet.getSheetView().setDisplayGridlines(false);

        //Set Value.
        worksheet.getRange("B1").setValue("To-Do List");
        worksheet.getRange("B2:I7").setValue(data);

        //Set formula.
        worksheet.getRange("E3").setFormula("=TODAY()");
        worksheet.getRange("E4").setFormula("=TODAY()-30");
        worksheet.getRange("E5").setFormula("=TODAY()-23");
        worksheet.getRange("E6").setFormula("=TODAY()-15");
        worksheet.getRange("E7").setFormula("=TODAY()-5");

        //Change the range's RowHeight and ColumnWidth.
        worksheet.setStandardHeight(30);
        worksheet.setStandardWidth(8.88671875);
        worksheet.getRange("1:1").setRowHeight(72.75);
        worksheet.getRange("2:2").setRowHeight(33);
        worksheet.getRange("3:7").setRowHeight(30.25);
        worksheet.getRange("A:A").setColumnWidth(2.77734375);
        worksheet.getRange("B:B").setColumnWidth(29.109375);
        worksheet.getRange("C:G").setColumnWidth(16.77734375);
        worksheet.getRange("H:H").setColumnWidth(10.77734375);
        worksheet.getRange("I:I").setColumnWidth(29.6640625);
        worksheet.getRange("J:J").setColumnWidth(2.77734375);

        //Modify the build in name getStyles().
        IStyle nameStyle_Normal = workbook.getStyles().get("Normal");
        nameStyle_Normal.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Normal.setWrapText(true);
        nameStyle_Normal.getFont().setThemeFont(ThemeFont.Minor);
        nameStyle_Normal.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Normal.getFont().setTintAndShade(0.25);

        IStyle nameStyle_Title = workbook.getStyles().get("Title");
        nameStyle_Title.setHorizontalAlignment(HorizontalAlignment.General);
        nameStyle_Title.setVerticalAlignment(VerticalAlignment.Bottom);
        nameStyle_Title.getFont().setThemeFont(ThemeFont.Minor);
        nameStyle_Title.getFont().setBold(true);
        nameStyle_Title.getFont().setSize(38);
        nameStyle_Title.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Title.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Title.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        nameStyle_Title.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Dark1);
        nameStyle_Title.setIncludeAlignment(true);
        nameStyle_Title.setIncludeBorder(true);

        IStyle nameStyle_Percent = workbook.getStyles().get("Percent");
        nameStyle_Percent.setHorizontalAlignment(HorizontalAlignment.Right);
        nameStyle_Percent.setIndentLevel(1);
        nameStyle_Percent.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Percent.setIncludeAlignment(true);

        IStyle nameStyle_Heading_1 = workbook.getStyles().get("Heading 1");
        nameStyle_Heading_1.setHorizontalAlignment(HorizontalAlignment.Left);
        nameStyle_Heading_1.setVerticalAlignment(VerticalAlignment.Bottom);
        nameStyle_Heading_1.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_1.getFont().setBold(false);
        nameStyle_Heading_1.getFont().setSize(11);
        nameStyle_Heading_1.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Heading_1.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Heading_1.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_1.setIncludeNumber(true);
        nameStyle_Heading_1.setIncludeAlignment(true);
        nameStyle_Heading_1.setIncludeBorder(false);

        IStyle nameStyle_Heading_2 = workbook.getStyles().get("Heading 2");
        nameStyle_Heading_2.setHorizontalAlignment(HorizontalAlignment.Right);
        nameStyle_Heading_2.setIndentLevel(2);
        nameStyle_Heading_2.setVerticalAlignment(VerticalAlignment.Bottom);
        nameStyle_Heading_2.getFont().setThemeFont(ThemeFont.Major);
        nameStyle_Heading_2.getFont().setBold(false);
        nameStyle_Heading_2.getFont().setSize(11);
        nameStyle_Heading_2.getFont().setThemeColor(ThemeColor.Dark2);
        nameStyle_Heading_2.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        nameStyle_Heading_2.setIncludeNumber(true);
        nameStyle_Heading_2.setIncludeAlignment(true);

        //Create custom name styes.
        IStyle nameStyle_Done = workbook.getStyles().add("Done");
        nameStyle_Done.setNumberFormat("\"Done\";\"\";\"\"");
        nameStyle_Done.setHorizontalAlignment(HorizontalAlignment.Center);
        nameStyle_Done.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Done.getFont().setThemeFont(ThemeFont.Minor);
        nameStyle_Done.getFont().setThemeColor(ThemeColor.Light1);

        IStyle nameStyle_Date = workbook.getStyles().add("Date");
        nameStyle_Date.setNumberFormat("yyyy/m/d");
        nameStyle_Date.setHorizontalAlignment(HorizontalAlignment.Right);
        nameStyle_Date.setVerticalAlignment(VerticalAlignment.Center);
        nameStyle_Date.getFont().setThemeFont(ThemeFont.Minor);
        nameStyle_Date.getFont().setThemeColor(ThemeColor.Dark1);
        nameStyle_Date.getFont().setTintAndShade(0.249946592608417);
        nameStyle_Date.setIncludeBorder(false);
        nameStyle_Date.setIncludePatterns(false);

        //Apply the above name styles on ranges.
        worksheet.getRange("B1:I1").setStyle(workbook.getStyles().get("Title"));
        worksheet.getRange("B2:D2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("E2:F2").setStyle(workbook.getStyles().get("Heading 2"));
        worksheet.getRange("G2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("H2:H7").setStyle(workbook.getStyles().get("Done"));
        worksheet.getRange("I2").setStyle(workbook.getStyles().get("Heading 1"));
        worksheet.getRange("E3:F7").setStyle(workbook.getStyles().get("Date"));
        worksheet.getRange("G3:G7").setStyle(workbook.getStyles().get("Percent"));

        //Add one custom table style.
        ITableStyle style = workbook.getTableStyles().add("To-do List");
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Light1);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setTintAndShade(-0.14993743705557422);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setThemeColor(ThemeColor.Light1);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setTintAndShade(-0.14993743705557422);

        //Create a table and apply the above table style.
        ITable table = worksheet.getTables().add(worksheet.getRange("B2:I7"), true);
        table.setName("ToDoList");
        table.setTableStyle(style);

        //Use table formula in table range.
        worksheet.getRange("F3").setFormula("=[@[START DATE]]+7");
        worksheet.getRange("F4").setFormula("=[@[START DATE]]+35");
        worksheet.getRange("F5").setFormula("=[@[START DATE]]+10");
        worksheet.getRange("F6").setFormula("=[@[START DATE]]+36");
        worksheet.getRange("F7").setFormula("=[@[START DATE]]+14");
        worksheet.getRange("H3:H7").setFormula("=--([@[% COMPLETE]]>=1)");

        //Add a expression rule.
        IFormatCondition expression = (IFormatCondition) worksheet.getRange("B3:I7").getFormatConditions().add(FormatConditionType.Expression, FormatConditionOperator.Between, "=AND($G3=0,$G3<>\"\")", null);
        expression.getInterior().setThemeColor(ThemeColor.Light1);
        expression.getInterior().setTintAndShade(-0.0499893185216834);

        //Add a data bar rule.
        IDataBar dataBar = worksheet.getRange("G3:G7").getFormatConditions().addDatabar();
        dataBar.setBarFillType(DataBarFillType.Solid);
        dataBar.getBarColor().setThemeColor(ThemeColor.Accent1);
        dataBar.getBarColor().setTintAndShade(0.39997558519241921);

        //Add an icon set rule.
        IIconSetCondition iconSet = worksheet.getRange("H3:H7").getFormatConditions().addIconSetCondition();
        iconSet.setIconSet(workbook.getIconSets().get(IconSetType.Icon3Symbols));
        iconSet.getIconCriteria().get(2).setOperator(FormatConditionOperator.GreaterEqual);
        iconSet.getIconCriteria().get(2).setValue(1);
        iconSet.getIconCriteria().get(2).setType(ConditionValueTypes.Number);
        iconSet.getIconCriteria().get(1).setOperator(FormatConditionOperator.GreaterEqual);
        iconSet.getIconCriteria().get(1).setValue(0);
        iconSet.getIconCriteria().get(1).setType(ConditionValueTypes.Number);

        //Add a cell value rule.
        IFormatCondition cellValue = (IFormatCondition) worksheet.getRange("H3:H7").getFormatConditions().add(FormatConditionType.CellValue, FormatConditionOperator.NotEqual, 1, null);
        cellValue.setStopIfTrue(true);

        //Create list validations.
        worksheet.getRange("C3:C7").getValidation().add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Low, Normal, High", null);
        IValidation validation = worksheet.getRange("C3:C7").getValidation();
        validation.setErrorMessage("Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection");

        worksheet.getRange("D3:D7").getValidation().add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "Not Started,In Progress, Deferred, Complete", null);
        validation = worksheet.getRange("D3:D7").getValidation();
        validation.setErrorMessage("Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection");

        worksheet.getRange("G3:G7").getValidation().add(ValidationType.List, ValidationAlertStyle.Warning, ValidationOperator.Between, "0%,25%,50%,75%,100%", null);
        validation = worksheet.getRange("G3:G7").getValidation();
        validation.setErrorMessage("Select entry from the list. Select CANCEL, then press ALT+DOWN ARROW to navigate the list. Select ENTER to make selection");

        //Create custom validation.
        worksheet.getRange("F3:F7").getValidation().add(ValidationType.Custom, ValidationAlertStyle.Warning, ValidationOperator.Between, "=F3>=E3", null);
        validation = worksheet.getRange("F3:F7").getValidation();
        validation.setErrorMessage("The Due Date must be greater than or equal to the Start Date. Select YES to keep the value, NO to retry or CANCEL to clear the entry");

        //Create none validations, set inputmessage.
        worksheet.getRange("B2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("B2").getValidation();
        validation.setInputMessage("Enter Task in this column under this heading. Use heading filters to find specific entries");

        worksheet.getRange("C2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("C2").getValidation();
        validation.setInputMessage("Select Priority in this column under this heading. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection");

        worksheet.getRange("D2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("D2").getValidation();
        validation.setInputMessage("Select Status in this column under this heading.  Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection");

        worksheet.getRange("E2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("E2").getValidation();
        validation.setInputMessage("Enter Start Date in this column under this heading");

        worksheet.getRange("F2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("F2").getValidation();
        validation.setInputMessage("Enter Due Date in this column under this heading");

        worksheet.getRange("G2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("G2").getValidation();
        validation.setInputMessage("Select % Complete in this column. Press ALT+DOWN ARROW to open the drop-down list, then ENTER to make selection. A status bar indicates progress toward completion");

        worksheet.getRange("H2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("H2").getValidation();
        validation.setInputMessage("Icon indicator for task completion in this column under this heading is automatically updated as tasks complete");

        worksheet.getRange("I2").getValidation().add(ValidationType.None, ValidationAlertStyle.Stop, ValidationOperator.Between, null, null);
        validation = worksheet.getRange("I2").getValidation();
        validation.setInputMessage("Enter Notes in this column under this heading");

        //Create customize theme.
        ITheme theme = new Theme("test");
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.FromArgb(0, 0, 0));
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.FromArgb(255, 255, 255));
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.FromArgb(37, 28, 34));
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.FromArgb(240, 248, 246));
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.FromArgb(148, 112, 135));
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.FromArgb(71, 166, 181));
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.FromArgb(234, 194, 53));
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.FromArgb(107, 192, 129));
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.FromArgb(233, 115, 61));
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.FromArgb(251, 147, 59));
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.FromArgb(71, 166, 181));
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.FromArgb(148, 112, 135));
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Helvetica Neue");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Bookman Old Style");

        //Apply the above custom theme.
        workbook.setTheme(theme);

        //Set active cell.
        worksheet.getRange("G4").activate();

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
