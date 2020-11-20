package com.grapecity.documents.excel.examples.features.theme;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Themes;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ChangeWorkbookTheme extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Change workbook's theme to builtin theme.
        workbook.setTheme(Themes.GetBerlin());

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B2").setValue("Major Font:");
        worksheet.getRange("B3").setValue("Minor Font:");
        worksheet.getRange("C2").setValue("Trebuchet MS");
        worksheet.getRange("C3").setValue("Trebuchet MS");
        worksheet.getRange("C2").getFont().setThemeFont(ThemeFont.Major);
        worksheet.getRange("C3").getFont().setThemeFont(ThemeFont.Minor);

        worksheet.getRange("E2:E13").setValue(new Object[]{
                "Light1", "Dark1", "Light2", "Dark2", "Accent1", "Accent2",
                "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"
        });

        worksheet.getRange("F2").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F3").getInterior().setThemeColor(ThemeColor.Dark1);
        worksheet.getRange("F4").getInterior().setThemeColor(ThemeColor.Light2);
        worksheet.getRange("F5").getInterior().setThemeColor(ThemeColor.Dark2);
        worksheet.getRange("F6").getInterior().setThemeColor(ThemeColor.Accent1);
        worksheet.getRange("F7").getInterior().setThemeColor(ThemeColor.Accent2);
        worksheet.getRange("F8").getInterior().setThemeColor(ThemeColor.Accent3);
        worksheet.getRange("F9").getInterior().setThemeColor(ThemeColor.Accent4);
        worksheet.getRange("F10").getInterior().setThemeColor(ThemeColor.Accent5);
        worksheet.getRange("F11").getInterior().setThemeColor(ThemeColor.Accent6);
        worksheet.getRange("F12").getInterior().setThemeColor(ThemeColor.Hyperlink);
        worksheet.getRange("F13").getInterior().setThemeColor(ThemeColor.FollowedHyperlink);

        worksheet.getRange("B:F").setColumnWidth(15);

    }

}
