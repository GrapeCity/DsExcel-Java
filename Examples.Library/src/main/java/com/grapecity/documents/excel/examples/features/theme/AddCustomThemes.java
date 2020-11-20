package com.grapecity.documents.excel.examples.features.theme;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddCustomThemes extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        // Base theme is office theme when not give parameter.
        ITheme theme = new Theme("testTheme");
        // ITheme theme = new GrapeCity.Documents.Excel.Theme("testTheme",Themes.Badge);

        // Customize theme's color.
        theme.getThemeColorScheme().get(ThemeColor.Light1).setRGB(Color.GetAntiqueWhite());
        theme.getThemeColorScheme().get(ThemeColor.Dark1).setRGB(Color.GetAqua());
        theme.getThemeColorScheme().get(ThemeColor.Light2).setRGB(Color.GetBeige());
        theme.getThemeColorScheme().get(ThemeColor.Dark2).setRGB(Color.GetBlack());
        theme.getThemeColorScheme().get(ThemeColor.Accent1).setRGB(Color.GetCadetBlue());
        theme.getThemeColorScheme().get(ThemeColor.Accent2).setRGB(Color.GetChartreuse());
        theme.getThemeColorScheme().get(ThemeColor.Accent3).setRGB(Color.GetChocolate());
        theme.getThemeColorScheme().get(ThemeColor.Accent4).setRGB(Color.GetCoral());
        theme.getThemeColorScheme().get(ThemeColor.Accent5).setRGB(Color.GetCornflowerBlue());
        theme.getThemeColorScheme().get(ThemeColor.Accent6).setRGB(Color.GetCornsilk());
        theme.getThemeColorScheme().get(ThemeColor.Hyperlink).setRGB(Color.GetHoneydew());
        theme.getThemeColorScheme().get(ThemeColor.FollowedHyperlink).setRGB(Color.GetHotPink());

        // Customize theme's font.
        theme.getThemeFontScheme().getMajor().get(FontLanguageIndex.Latin).setName("Kristen ITC");
        theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Segoe Script");

        // Change workbook's theme to custom theme.
        workbook.setTheme(theme);

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("B2").setValue("Major Font:");
        worksheet.getRange("B3").setValue("Minor Font:");
        worksheet.getRange("C2").setValue("Kristen ITC");
        worksheet.getRange("C3").setValue("Segoe Script");
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

    }
}
