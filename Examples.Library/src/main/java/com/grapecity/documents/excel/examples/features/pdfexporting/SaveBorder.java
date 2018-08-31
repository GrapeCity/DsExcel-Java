package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveBorder extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Single cell border
        sheet.getRange("B2").getBorders().setThemeColor(ThemeColor.Accent1);
        sheet.getRange("B2").getBorders().setLineStyle(BorderLineStyle.SlantDashDot);
        sheet.getRange("B2").getBorders().get(BordersIndex.DiagonalUp).setThemeColor(ThemeColor.Accent1);
        sheet.getRange("B2").getBorders().get(BordersIndex.DiagonalUp).setLineStyle(BorderLineStyle.SlantDashDot);
        sheet.getRange("B2").getBorders().get(BordersIndex.DiagonalDown).setThemeColor(ThemeColor.Accent1);
        sheet.getRange("B2").getBorders().get(BordersIndex.DiagonalDown).setLineStyle(BorderLineStyle.SlantDashDot);

        //Range border
        sheet.getRange("D2:E3").getBorders().setThemeColor(ThemeColor.Accent1);
        sheet.getRange("D2:E3").getBorders().setLineStyle(BorderLineStyle.DashDot);
        sheet.getRange("D2:E3").getBorders().get(BordersIndex.DiagonalDown).setThemeColor(ThemeColor.Accent1);
        sheet.getRange("D2:E3").getBorders().get(BordersIndex.DiagonalDown).setLineStyle(BorderLineStyle.DashDot);

        //Merge cell border
        sheet.getRange("B6:C7").merge();
        sheet.getRange("B6:C7").getBorders().setThemeColor(ThemeColor.Accent1);
        sheet.getRange("B6:C7").getBorders().setLineStyle(BorderLineStyle.Double);
        sheet.getRange("B6:C7").getBorders().get(BordersIndex.DiagonalUp).setThemeColor(ThemeColor.Accent1);
        sheet.getRange("B6:C7").getBorders().get(BordersIndex.DiagonalUp).setLineStyle(BorderLineStyle.Double);

        //Border style on table
        ITable table = sheet.getTables().add(sheet.getRange("B12:G22"), true);

        //Create custom table style
        ITableStyle customTableStyle = workbook.getTableStyles().get("TableStyleMedium10").duplicate();

        //Set outline border for "whole table" style
        ITableStyleElement wholeTableStyle = customTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeTop).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeRight).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeLeft).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thick);

        //Set vertical border for "first row strip" style
        ITableStyleElement firstRowStripStyle = customTableStyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe);
        firstRowStripStyle.getBorders().get(BordersIndex.InsideVertical).setThemeColor(ThemeColor.Accent6);
        firstRowStripStyle.getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Dashed);

        //Apply custom style to table
        table.setTableStyle(customTableStyle);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}