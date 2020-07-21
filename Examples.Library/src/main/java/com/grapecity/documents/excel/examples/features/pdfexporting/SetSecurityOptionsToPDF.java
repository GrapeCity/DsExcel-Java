package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PdfSaveOptions;
import com.grapecity.documents.excel.PdfSecurityOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetSecurityOptionsToPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue("GrapeCity Documents for Excel");
        worksheet.getRange("A1").getFont().setSize(25);
        
        //The security settings of pdf when converting excel to pdf.
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        //Sets the user password.
        securityOptions.setUserPassword("user");
        //Sets the owner password.
        securityOptions.setOwnerPassword("owner");
        //Allow to print pdf document.
        securityOptions.setPrintPermission(true);
        //Print the pdf document in high quality.
        securityOptions.setFullQualityPrintPermission(true);
        //Allow to copy or extract the content of the pdf document.
        securityOptions.setExtractContentPermission(true);
        //Allow to modify the pdf document.
        securityOptions.setModifyDocumentPermission(true);
        //Allow to insert, rotate, or delete pages and create bookmarks or thumbnail images of the pdf document.
        securityOptions.setAssembleDocumentPermission(true);
        //Allow to modify text annotations and fill the form fields of the pdf document.
        securityOptions.setModifyAnnotationsPermission(true);
        //Filling the form fields of the pdf document is not allowed.
        securityOptions.setFillFormsPermission(false);


        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        //Sets the secutity settings of the pdf.
        pdfSaveOptions.setSecurityOptions(securityOptions);

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
