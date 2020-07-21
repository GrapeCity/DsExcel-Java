package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.ByteArrayOutputStream;
import java.util.GregorianCalendar;

import com.grapecity.documents.excel.DocumentProperties;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PdfSaveOptions;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetDocumentPropertiesToPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").setValue("GrapeCity Documents for Excel");
        worksheet.getRange("A1").getFont().setSize(25);
        
        DocumentProperties documentProperties = new DocumentProperties();
        //Sets the name of the person that created the PDF document.
        documentProperties.setAuthor("Jaime Smith");
        //Sets the title of the  PDF document.
        documentProperties.setTitle("GcPdf Document Info Sample");
        //Set the PDF version.
        documentProperties.setPdfVersion(1.5f);
        //Set the subject of the PDF document.
        documentProperties.setSubject("GcPdfDocument.DocumentInfo");
        //Set the keyword associated with the PDF document.
        documentProperties.setKeywords("Keyword1");
        //Set the creation date and time of the PDF document.
        documentProperties.setCreationDate(new GregorianCalendar(2019,5,24));
        //Set the date and time the PDF document was most recently modified.
        documentProperties.setModifyDate(new GregorianCalendar(2020,5,24));
        //Set the name of the application that created the original PDF document.
        documentProperties.setCreator("GcPdfWeb Creator");
        //Set the name of the application that created the PDF document.
        documentProperties.setProducer("GcPdfWeb Producer");


        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        //Sets the document properties of the pdf.
        pdfSaveOptions.setDocumentProperties(documentProperties);

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
