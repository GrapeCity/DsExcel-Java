package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CenterAcrossSelection  extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {
        IWorksheet activeSheet = workbook.getActiveSheet();

        // Fill data
        Object[][] testData = new Object[][] {
            {"Fruit", null, null, "Fruit Total", "Vegetables", null, null, null, "Vegetables Total", "Grand Total"},
            {"Canada", "Germany", "New Zealand", null, "Australia", "Germany", "United Kingdom", "United States", null, null},
            {7431d, null, null, 9848d, null, null, null, null, null, 9848d},
            {8384d, 8250d, 6906d, 24157d, null, null, null, null, null, 24157d},
            {null, null, null, null, null, 2626d, null, null, 2626d, 2626d},
            {null, null, null, null, 9062d, null, 8239d, 7012d, 24313d, 24313d},
            {null, null, null, null, null, 1903d, null, 4270d, 6173d, 6173d},
            {null, null, null, 3610d, null, null, null, null, null, 3610d},
            {15815d, 8250d, 6906d, 30971d, 9062d, 4529d, 8239d, 11282d, 33112d, 64083d}
        };
        IRange dataAndHeader = activeSheet.getRange("$A$1:$J$9");
        dataAndHeader.setValue(testData);

        // Add borders
        IBorders borders = dataAndHeader.getBorders();
        borders.setLineStyle(BorderLineStyle.Thin);
        borders.setThemeColor(ThemeColor.Dark1);
        borders.setTintAndShade(-0.499984740745262);

        // Use center across selection in the header row
        IRange header = activeSheet.getRange("$A$1:$J$1");
        header.setHorizontalAlignment(HorizontalAlignment.CenterContinuous);

        activeSheet.getColumns().autoFit();
        activeSheet.getPageSetup().setOrientation(PageOrientation.Landscape);
	}

    @Override
    public boolean getSavePdf()
    {
        return true;
    }

    @Override
    public boolean getIsNew()
    {
        return true;
    }

    @Override
    public boolean getShowViewer()
    {
        return false;
    }

}