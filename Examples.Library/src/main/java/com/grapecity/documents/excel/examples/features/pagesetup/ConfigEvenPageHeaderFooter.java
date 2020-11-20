package com.grapecity.documents.excel.examples.features.pagesetup;

import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigEvenPageHeaderFooter extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        workbook.open(this.getResourceStream("xlsx/PageSetup Demo.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set even page headerfooter
        worksheet.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        worksheet.getPageSetup().getEvenPage().getCenterHeader().setText("&T");
        worksheet.getPageSetup().getEvenPage().getRightFooter().setText("&D");

        //Set even page headerfooter's graphic
        worksheet.getPageSetup().getEvenPage().getLeftFooter().setText("&G");
        InputStream stream = this.getResourceStream("logo.png");
        worksheet.getPageSetup().getEvenPage().getLeftFooter().getPicture().setGraphicStream(stream, ImageType.PNG);

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
