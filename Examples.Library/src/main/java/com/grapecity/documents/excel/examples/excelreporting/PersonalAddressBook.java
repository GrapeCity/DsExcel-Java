package com.grapecity.documents.excel.examples.excelreporting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Themes;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PersonalAddressBook extends ExampleBase {

    @Override
    public void beforeExecute(Workbook workbook, String userAgents) {

        if (userAgents.toLowerCase().contains("macintosh")) {

            ITheme theme = new Theme("testTheme", Themes.GetOfficeTheme());
            theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Trebuchet MS");
            workbook.setTheme(theme);
            IStyle style_Normal = workbook.getStyles().get("Normal");
            style_Normal.getFont().setThemeFont(ThemeFont.Minor);
        }
    }

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // ***************************Set RowHeight & Width****************************
        worksheet.setStandardHeight(30);
        worksheet.getRange("3:4").setRowHeight(30.25);
        worksheet.getRange("1:1").setRowHeight(103.50);
        worksheet.getRange("2:2").setRowHeight(38.25);
        worksheet.getRange("A:A").setColumnWidth(2.625);
        worksheet.getRange("B:B").setColumnWidth(22.25);
        worksheet.getRange("C:E").setColumnWidth(17.25);
        worksheet.getRange("F:F").setColumnWidth(31.875);
        worksheet.getRange("G:G").setColumnWidth(22.625);
        worksheet.getRange("H:H").setColumnWidth(30);
        worksheet.getRange("I:I").setColumnWidth(20.25);
        worksheet.getRange("J:J").setColumnWidth(17.625);
        worksheet.getRange("K:K").setColumnWidth(12.625);
        worksheet.getRange("L:L").setColumnWidth(37.25);
        worksheet.getRange("M:M").setColumnWidth(2.625);

        // *******************************Set Table Value &
        // Formulas*************************************
        ITable table = worksheet.getTables().add(worksheet.getRange("B2:L4"), true);
        worksheet.getRange("B2:L4").setValue(new Object[][]{
                {"NAME", "WORK", "CELL", "HOME", "EMAIL", "BIRTHDAY", "ADDRESS", "CITY", "STATE", "ZIP", "NOTE"},
                {"Kim Abercrombie", 1235550123, 1235550123, 1235550123, "someone@example.com", null, "123 N. Maple",
                        "Cherryville", "WA", 98031, ""},
                {"John Smith", 3215550123L, "", "", "someone@example.com", null, "456 E. Aspen", "", "", "", ""},});
        worksheet.getRange("B1").setValue("ADDRESS BOOK");
        worksheet.getRange("G3").setFormula("=TODAY()");
        worksheet.getRange("G4").setFormula("=TODAY()+5");

        // ****************************Set Table Style********************************
        ITableStyle tableStyle = workbook.getTableStyles().add("Personal Address Book");
        workbook.setDefaultTableStyle("Personal Address Book");

        // Set WholeTable element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setColor(Color.FromArgb(179, 35, 23));
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Thin);
        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);

        // Set FirstColumn element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.FirstColumn).getFont().setBold(true);

        // Set SecondColumns element style.
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().setColor(Color.FromArgb(179, 35, 23));
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thick);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);

        // ***********************************Set Named
        // Styles*****************************
        IStyle normalStyle = workbook.getStyles().get("Normal");
        normalStyle.getFont().setName("Arial");
        normalStyle.getFont().setColor(Color.FromArgb(179, 35, 23));
        normalStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        normalStyle.setIndentLevel(1);
        normalStyle.setVerticalAlignment(VerticalAlignment.Center);
        normalStyle.setWrapText(true);

        IStyle titleStyle = workbook.getStyles().get("Title");
        titleStyle.setIncludeAlignment(true);
        titleStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        titleStyle.setVerticalAlignment(VerticalAlignment.Center);
        titleStyle.getFont().setName("Arial");
        titleStyle.getFont().setBold(true);
        titleStyle.getFont().setSize(72);
        titleStyle.getFont().setColor(Color.FromArgb(179, 35, 23));

        IStyle heading1Style = workbook.getStyles().get("Heading 1");
        heading1Style.setIncludeBorder(false);
        heading1Style.getFont().setName("Arial");
        heading1Style.getFont().setSize(18);
        heading1Style.getFont().setColor(Color.FromArgb(179, 35, 23));

        IStyle dataStyle = workbook.getStyles().add("Data");
        dataStyle.setIncludeNumber(true);
        dataStyle.setNumberFormat("m/d/yyyy");

        IStyle phoneStyle = workbook.getStyles().add("Phone");
        phoneStyle.setIncludeNumber(true);
        phoneStyle.setNumberFormat("[<=9999999]###-####;(###) ###-####");

        // ****************************************Use
        // NamedStyle**************************
        worksheet.getSheetView().setDisplayGridlines(false);
        worksheet.getRange("B2:L2").getInterior().setColor(Color.FromArgb(217, 217, 217));
        worksheet.getRange("B3:B4").getFont().setBold(true);
        worksheet.getRange("2:2").setHorizontalAlignment(HorizontalAlignment.Left);

        table.setTableStyle(tableStyle);
        worksheet.getRange("B1").setStyle(titleStyle);
        worksheet.getRange("B2:L2").setStyle(heading1Style);
        worksheet.getRange("C3:E4").setStyle(phoneStyle);
        worksheet.getRange("G3:G4").setStyle(dataStyle);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
