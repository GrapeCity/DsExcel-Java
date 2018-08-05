package com.grapecity.documents.excel.examples.features.hyperlinks;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class DeleteHyperlinks extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:A").setColumnWidth(30);

        //add a hyperlink link to web page.
        worksheet.getRange("A1:B2").getHyperlinks().add(worksheet.getRange("A1"), "http://www.google.com/", null, "open google web site.", "Google");

        //add a hyperlink link to a range in this document.
        worksheet.getRange("A3:B4").getHyperlinks().add(worksheet.getRange("A3"), null, "Sheet1!$C$3:$E$4", "Go to sheet1 C3:E4", "");

        //add a hyperlink link to email address.
        worksheet.getRange("A5:B6").getHyperlinks().add(worksheet.getRange("A5"), "mailto:us.sales@grapecity.com", null, "Send an email to sales", "Send an email to sales");

        //add a hyperlink link to external file.
        //change the path to real picture file path.
        String path = "external.xlsx";
        worksheet.getRange("A7:B8").getHyperlinks().add(worksheet.getRange("A7"), path, null, "link to external.xlsx file.", "External.xlsx");

        //delete hyperlinks in range A1:A2.
        worksheet.getRange("A1:A2").getHyperlinks().delete();

        //delete all hyperlinks in this worksheet.
        worksheet.getHyperlinks().delete();

    }

}
