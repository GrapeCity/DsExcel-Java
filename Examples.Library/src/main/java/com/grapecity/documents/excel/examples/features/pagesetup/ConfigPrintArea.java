package com.grapecity.documents.excel.examples.features.pagesetup;

import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigPrintArea extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/PageSetup Demo.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set print area & print titles
        worksheet.getPageSetup().setPrintArea("$D$5:$G$10");
        worksheet.getPageSetup().setPrintTitleRows("$5:$10");
        worksheet.getPageSetup().setPrintTitleColumns("$D:$G");

    }

    @Override
    public String getTemplateName() {
        return "PageSetup Demo.xlsx";
    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/PageSetup Demo.xlsx"};
    }
}
