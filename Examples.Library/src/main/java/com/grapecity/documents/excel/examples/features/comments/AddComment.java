package com.grapecity.documents.excel.examples.features.comments;

import com.grapecity.documents.excel.IComment;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddComment extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //create comment for range C3.
        IComment comment = worksheet.getRange("C3").addComment("Range C3's comment.");

        //change comment's text.
        comment.setText("Range C3's new comment");

    }

}
