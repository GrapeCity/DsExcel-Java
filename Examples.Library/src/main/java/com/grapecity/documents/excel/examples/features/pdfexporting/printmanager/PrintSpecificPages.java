package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PrintSpecificPages extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/PrintSpecificPDFPages.xlsx"));

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Get the natural pagination information of the workbook.
        List<PageInfo> pages = printManager.paginate(workbook);

        //Pick some pages to print.
        List<PageInfo> newPages = new ArrayList<PageInfo>();
        newPages.add(pages.get(0));
        newPages.add(pages.get(2));

        //Update the page number and the page settings of each page. The page number is continuous.
        printManager.updatePageNumberAndPageSettings(newPages);

        //Save the pages into pdf file.
        printManager.savePageInfosToPDF(outputStream, newPages);
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
		return "PrintSpecificPDFPages.xlsx";
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/PrintSpecificPDFPages.xlsx" };
	}

}
