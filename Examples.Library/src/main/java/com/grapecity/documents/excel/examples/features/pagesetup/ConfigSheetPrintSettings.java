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

        URL url = ClassLoader.getSystemResource("xlsx/PageSetup Demo.xlsx");
        String filePath = url.getPath().substring(1).replaceAll("%20", " ");
        workbook.open(filePath);
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set sheet
        worksheet.getPageSetup().setPrintGridlines(true);
        worksheet.getPageSetup().setPrintHeadings(true);
        worksheet.getPageSetup().setBlackAndWhite(true);
        worksheet.getPageSetup().setPrintComments(PrintLocation.InPlace);
        worksheet.getPageSetup().setPrintErrors(PrintErrors.Dash);

    }

}
