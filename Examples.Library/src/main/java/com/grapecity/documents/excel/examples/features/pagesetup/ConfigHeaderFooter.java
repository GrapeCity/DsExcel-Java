package com.grapecity.documents.excel.examples.features.pagesetup;

import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigHeaderFooter extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/PageSetup Demo.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set page headerfooter
        worksheet.getPageSetup().setLeftHeader("&\"Arial,Italic\"LeftHeader");
        worksheet.getPageSetup().setCenterHeader("&P");

        //Set page headerfooter's graphic
        worksheet.getPageSetup().setCenterFooter("&G");
        InputStream stream = this.getResourceStream("logo.png");
        worksheet.getPageSetup().getCenterFooterPicture().setGraphicStream(stream, ImageType.PNG);

        //If you have picture resources locally, you can also set graphic in this way.
        //worksheet.getPageSetup().setCenterFooter("&G");
        //worksheet.getPageSetup().getCenterFooterPicture().setFilename("C:\\picture.png");
        //worksheet.PageSetup.CenterFooterPicture.Filename = @"C:\picture.png";

    }

    @Override
    public String getTemplateName() {
        return "PageSetup Demo.xlsx";
    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png", "xlsx/PageSetup Demo.xlsx"};
    }

}
