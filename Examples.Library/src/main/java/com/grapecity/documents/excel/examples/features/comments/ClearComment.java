package com.grapecity.documents.excel.examples.features.comments;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ClearComment extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("C3").addComment("Range C3's comment.");
        worksheet.getRange("D4").addComment("Range D4's comment.");
        worksheet.getRange("D5").addComment("Range D5's comment.");

        // delete a single cell comment.
        worksheet.getRange("D5").getComment().delete();

        // clear a range of cells comment.
        worksheet.getRange("C3:D4").clearComments();

    }

}
