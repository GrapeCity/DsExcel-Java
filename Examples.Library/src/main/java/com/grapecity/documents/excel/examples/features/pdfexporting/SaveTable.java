package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveTable extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Add Table
        ITable table = sheet.getTables().add(sheet.getRange("B5:G16"), true);
        table.setShowTotals(true);

        //Set values
        int[] data = new int[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11};
        sheet.getRange("C6:C16").setValue(data);
        sheet.getRange("D6:D16").setValue(data);

        //Set total functions
        table.getColumns().get(1).setTotalsCalculation(TotalsCalculation.Average);
        table.getColumns().get(2).setTotalsCalculation(TotalsCalculation.Sum);

        //Create custom table style
        ITableStyle customTableStyle = workbook.getTableStyles().get("TableStyleMedium10").duplicate();

        ITableStyleElement wholeTableStyle = customTableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable);
        wholeTableStyle.getFont().setItalic(true);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeTop).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeRight).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeBottom).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeLeft).setThemeColor(ThemeColor.Accent1);
        wholeTableStyle.getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thick);

        ITableStyleElement firstRowStripStyle = customTableStyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe);
        firstRowStripStyle.getFont().setBold(true);

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