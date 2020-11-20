package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.util.List;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CustomPageInfos extends ExampleBase {
	 @Override
	 public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/KeepTogether.xlsx"));
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Firstly, create a printManager.
        PrintManager printManager = new PrintManager();

        //Get the natural pagination information of the worksheet.
        //The first page of the natural pagination is "A1:F37", the second page is from row "A38:F73" 
        List<PageInfo> pages = printManager.paginate(worksheet);

        //Custom the pageInfos.
        pages.get(0).getPageContent().setRange(worksheet.getRange("A1:F36")); // The first page is "A1:F36".
        pages.get(0).getPageSettings().setCenterHeader("&KFF0000&20 Budget summary report"); // The center header of the first page will show the text "Budget summary report".
        pages.get(0).getPageSettings().setCenterFooter("&KFF0000&20 Page &P"); // The center footer of the first page will show the page number "1".
        pages.get(1).getPageContent().setRange(worksheet.getRange("A37:F73"));// The second page is "A37:F73".

        //Save the modified pages into pdf file.
        printManager.savePageInfosToPDF(outputStream, pages);
	 }

	@Override
	public boolean getSavePageInfos() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

	@Override
	public String getTemplateName() {
        return "KeepTogether.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/KeepTogether.xlsx" };
	}

}
