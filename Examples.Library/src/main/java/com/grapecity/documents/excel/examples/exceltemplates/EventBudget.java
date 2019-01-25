package com.grapecity.documents.excel.examples.exceltemplates;

import java.net.URL;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class EventBudget extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Load template file Event budget.xlsx from resource
        workbook.open(this.getResourceStream("xlsx/Event budget.xlsx"));

        IWorksheet worksheet = workbook.getActiveSheet();

        //change range B2's font size.
        worksheet.getRange("B2").getFont().setSize(22);

        //change range E4's font style to bold.
        worksheet.getRange("E4").getFont().setBold(true);

        //change table style.
        worksheet.getTables().get("tblAdmissions").setTableStyle(workbook.getTableStyles().get("TableStyleLight10"));
        worksheet.getTables().get("tblAds").setTableStyle(workbook.getTableStyles().get("TableStyleLight10"));
        worksheet.getTables().get("tblVendors").setTableStyle(workbook.getTableStyles().get("TableStyleLight10"));
        worksheet.getTables().get("tblItems").setTableStyle(workbook.getTableStyles().get("TableStyleLight10"));

        //modify range F4:G5's cell style.
        worksheet.getRange("F4:G5").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F4:G5").getInterior().setTintAndShade(-0.15);
        worksheet.getRange("F4:G5").getFont().setThemeFont(ThemeFont.Major);
        worksheet.getRange("F4:G5").getFont().setSize(12);
        worksheet.getRange("F4:G5").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.None);
        worksheet.getRange("F5:G5").setNumberFormat("$#,##0.00");

        //modify table columns' style.
        worksheet.getRange("F8:G11, F15:G18, F22:G25, F29:G33").getInterior().setThemeColor(ThemeColor.Light1);
        worksheet.getRange("F8:G11, F15:G18, F22:G25, F29:G33").getInterior().setTintAndShade(-0.15);
        worksheet.getRange("E8:G11, E15:G18, E22:G25, E29:G33").setNumberFormat("$#,##0.00");

    }

    @Override
    public String getTemplateName() {

        return "Event budget.xlsx";

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Event budget.xlsx"};
    }
}
