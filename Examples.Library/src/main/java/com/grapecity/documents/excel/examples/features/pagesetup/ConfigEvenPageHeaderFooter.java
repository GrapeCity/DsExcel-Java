package com.grapecity.documents.excel.examples.features.pagesetup;

import java.io.InputStream;
import java.net.URL;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigEvenPageHeaderFooter extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        URL url = ClassLoader.getSystemResource("xlsx/PageSetup Demo.xlsx");
        String filePath = url.getPath().substring(1).replaceAll("%20", " ");
        System.out.println(filePath);
        workbook.open(filePath);
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set even page headerfooter
        worksheet.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        worksheet.getPageSetup().getEvenPage().getCenterHeader().setText("&T");
        worksheet.getPageSetup().getEvenPage().getRightFooter().setText("&D");

        //Set even page headerfooter's graphic
        worksheet.getPageSetup().getEvenPage().getLeftFooter().setText("&G");
        InputStream stream = ClassLoader.getSystemResourceAsStream("logo.png");
        worksheet.getPageSetup().getEvenPage().getLeftFooter().getPicture().setGraphicStream(stream, ImageType.PNG);

    }

}
