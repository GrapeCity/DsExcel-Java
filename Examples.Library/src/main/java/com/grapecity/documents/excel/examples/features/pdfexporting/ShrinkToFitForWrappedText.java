package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PdfSaveOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShrinkToFitForWrappedText extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getPageSetup().setPrintGridlines(true); 
        worksheet.getPageSetup().setPrintHeadings(true);

        //"A1" is ordinary wrapped text.
        worksheet.getRange("A1").setWrapText(true);
        worksheet.getRange("A1").setValue("GrapeCity Documents for Excel");
        worksheet.getRange("A1").setRowHeight(38);
        worksheet.getRange("A1").setColumnWidth(9);

        //The wrapped text "A2" will be shrink to fit.
        //worksheet.Range["A2"].Interior.Color = Color.LightGreen;
        worksheet.getRange("A2").setWrapText(true);
        worksheet.getRange("A2").setShrinkToFit(true);
        worksheet.getRange("A2").setValue("GrapeCity Documents for Excel");
        worksheet.getRange("A2").setRowHeight(38);

        //You must create a pdfSaveOptions object before using ShrinkToFitSettings.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        //Shrink the wrapped text within the cell with existing row height/column width, while exporting to PDF. 
        pdfSaveOptions.getShrinkToFitSettings().setCanShrinkToFitWrappedText(true);

//        //Set minimum font size with which the text should shrink.
//        pdfSaveOptions.getShrinkToFitSettings().setMinimumFont(10);
//        //If after setting the minimum font size, the text is very long not fully visible, the ellipsis string to show for long text.
//        pdfSaveOptions.getShrinkToFitSettings().setEllipsis("~");

        //Save the workbook into pdf file.
        workbook.save(outputStream, pdfSaveOptions);
    }

    @Override
    public boolean getSavePageInfos() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

}
