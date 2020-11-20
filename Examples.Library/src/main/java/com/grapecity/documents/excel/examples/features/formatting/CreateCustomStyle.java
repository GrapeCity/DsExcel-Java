package com.grapecity.documents.excel.examples.features.formatting;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Pattern;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.UnderlineType;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CreateCustomStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        // Add custom name style.
        IStyle style = workbook.getStyles().add("testStyle");

        // Config custom name style settings begin.
        // Border
        style.getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.Thin);
        style.getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thick);
        style.getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.Double);
        style.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Double);
        style.getBorders().setColor(Color.FromArgb(0, 255, 0));

        // Font
        style.getFont().setThemeColor(ThemeColor.Accent1);
        style.getFont().setTintAndShade(0.8);
        style.getFont().setItalic(true);
        style.getFont().setBold(true);
        style.getFont().setName("LiSu");
        style.getFont().setSize(28);
        style.getFont().setStrikethrough(true);
        style.getFont().setSubscript(true);
        style.getFont().setSuperscript(false);
        style.getFont().setUnderline(UnderlineType.Double);

        // Protection
        style.setFormulaHidden(true);
        style.setLocked(false);

        // Number
        style.setNumberFormat("#,##0_);[Red](#,##0)");

        // Alignment
        style.setHorizontalAlignment(HorizontalAlignment.Right);
        style.setVerticalAlignment(VerticalAlignment.Bottom);
        style.setWrapText(true);
        style.setIndentLevel(5);
        style.setOrientation(45);

        // Fill
        style.getInterior().setColorIndex(5);
        style.getInterior().setPattern(Pattern.Down);
        style.getInterior().setPatternColor(Color.FromArgb(0, 0, 255));

        style.setIncludeAlignment(false);
        style.setIncludeBorder(true);
        style.setIncludeFont(false);
        style.setIncludeNumber(true);
        style.setIncludePatterns(false);
        style.setIncludeProtection(true);
        // Config custom name style settings end.

        // Set range's style to custom name style.
        worksheet.getRange("A1").setValue("My test style");
        worksheet.getRange("A1").setStyle(worksheet.getWorkbook().getStyles().get("testStyle"));

    }

}
