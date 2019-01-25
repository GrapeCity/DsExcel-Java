package com.grapecity.documents.excel.examples.features.pagesetup;

import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigPaperScaling extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/PageSetup Demo.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set paper scaling
        //Method 1: Set percent scale 
        worksheet.getPageSetup().setIsPercentScale(true);
        worksheet.getPageSetup().setZoom(150);

        //Or Method 2: Fit to page's wide & tall
        //worksheet.getPageSetup().setIsPercentScale(false);
        //worksheet.getPageSetup().setFitToPagesWide(3);
        //worksheet.getPageSetup().setFitToPagesTall(4);

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
