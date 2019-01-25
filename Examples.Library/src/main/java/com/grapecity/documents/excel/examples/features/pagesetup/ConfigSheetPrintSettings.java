package com.grapecity.documents.excel.examples.features.pagesetup;

import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PrintErrors;
import com.grapecity.documents.excel.PrintLocation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigSheetPrintSettings extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/PageSetup Demo.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set sheet
        worksheet.getPageSetup().setPrintGridlines(true);
        worksheet.getPageSetup().setPrintHeadings(true);
        worksheet.getPageSetup().setBlackAndWhite(true);
        worksheet.getPageSetup().setPrintComments(PrintLocation.InPlace);
        worksheet.getPageSetup().setPrintErrors(PrintErrors.Dash);

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
