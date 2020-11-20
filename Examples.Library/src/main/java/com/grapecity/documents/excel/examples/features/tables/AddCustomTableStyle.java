package com.grapecity.documents.excel.examples.features.tables;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddCustomTableStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Add one custom table style.
        ITableStyle style = workbook.getTableStyles().add("test");

        //Set WholeTable element style.
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setItalic(true);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setColor(Color.GetWhite());
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getFont().setStrikethrough(true);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setLineStyle(BorderLineStyle.Dotted);
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getBorders().setColor(Color.FromArgb(0, 193, 213));
        style.getTableStyleElements().get(TableStyleElementType.WholeTable).getInterior().setColor(Color.FromArgb(59, 92, 170));

        //Set FirstColumnStripe element style.
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).getFont().setBold(true);
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).getFont().setColor(Color.FromArgb(255, 0, 0));
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).getBorders().setLineStyle(BorderLineStyle.Thick);
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).getBorders().setThemeColor(ThemeColor.Accent5);
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).getInterior().setColor(Color.FromArgb(255, 255, 0));
        style.getTableStyleElements().get(TableStyleElementType.FirstColumnStripe).setStripeSize(2);

        //Set SecondColumnStripe element style.
        style.getTableStyleElements().get(TableStyleElementType.SecondColumnStripe).getFont().setColor(Color.FromArgb(255, 0, 255));
        style.getTableStyleElements().get(TableStyleElementType.SecondColumnStripe).getBorders().setLineStyle(BorderLineStyle.DashDot);
        style.getTableStyleElements().get(TableStyleElementType.SecondColumnStripe).getBorders().setColor(Color.FromArgb(42, 105, 162));
        style.getTableStyleElements().get(TableStyleElementType.SecondColumnStripe).getInterior().setColor(Color.FromArgb(204, 204, 255));

        //add table.
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        ITable table = worksheet.getTables().add(worksheet.getRange("A1:F7"), true);
        worksheet.getRange("A:F").setColumnWidth(15);

        //set custom table style to table.
        table.setTableStyle(style);

    }

}
