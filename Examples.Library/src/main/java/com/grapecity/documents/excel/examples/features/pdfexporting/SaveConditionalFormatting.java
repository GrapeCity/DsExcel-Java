package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveConditionalFormatting extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Conditional formatting on merge cell
        sheet.getRange("B2:C4").merge();
        sheet.getRange("B2:C4").setValue(123);
        IFormatCondition cf = (IFormatCondition) sheet.getRange("B2:C4").getFormatConditions().add(FormatConditionType.CellValue, FormatConditionOperator.Greater, 0, 0);
        cf.getBorders().setThemeColor(ThemeColor.Accent1);
        cf.getBorders().setLineStyle(BorderLineStyle.Thin);

        //Set cell values
        int[] data = new int[]{1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
        sheet.getRange("B10:B19").setValue(data);
        sheet.getRange("C10:C19").setValue(data);
        sheet.getRange("D10:D19").setValue(data);

        //Set conditional formattings
        //Color scale
        IColorScale cf1 = sheet.getRange("B10:B19").getFormatConditions().addColorScale(ColorScaleType.ThreeColorScale);
        cf1.getColorScaleCriteria().get(0).setType(ConditionValueTypes.LowestValue);
        cf1.getColorScaleCriteria().get(0).getFormatColor().setColor(Color.FromRGB(248, 105, 107));
        cf1.getColorScaleCriteria().get(1).setType(ConditionValueTypes.Percentile);
        cf1.getColorScaleCriteria().get(1).setValue(50);
        cf1.getColorScaleCriteria().get(1).getFormatColor().setColor(Color.FromRGB(255, 235, 132));
        cf1.getColorScaleCriteria().get(2).setType(ConditionValueTypes.HighestValue);
        cf1.getColorScaleCriteria().get(2).getFormatColor().setColor(Color.FromRGB(99, 190, 123));

        //Data bar
        sheet.getRange("C14").setValue(-5);
        sheet.getRange("C17").setValue(-8);
        IDataBar cf2 = sheet.getRange("C10:C19").getFormatConditions().addDatabar();
        cf2.getMinPoint().setType(ConditionValueTypes.AutomaticMin);
        cf2.getMaxPoint().setType(ConditionValueTypes.AutomaticMax);
        cf2.setBarFillType(DataBarFillType.Gradient);
        cf2.getBarColor().setColor(Color.FromRGB(0, 138, 239));
        cf2.getBarBorder().getColor().setColor(Color.FromRGB(0, 138, 239));
        cf2.getNegativeBarFormat().getColor().setColor(Color.FromRGB(255, 0, 0));
        cf2.getNegativeBarFormat().setBorderColorType(DataBarNegativeColorType.Color);
        cf2.getNegativeBarFormat().getBorderColor().setColor(Color.FromRGB(255, 0, 0));
        cf2.getAxisColor().setColor(Color.getBlack());
        cf2.setAxisPosition(DataBarAxisPosition.Automatic);

        //Icon set
        IIconSetCondition cf3 = sheet.getRange("D10:D19").getFormatConditions().addIconSetCondition();
        cf3.setIconSet(workbook.getIconSets().get(IconSetType.Icon3Symbols));
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