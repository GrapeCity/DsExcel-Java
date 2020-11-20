package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class KeepTogether extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/KeepTogether.xlsx"));
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //The first page of the natural pagination is from row 1th to 37th, the second page is from row 38th to 73th. 
        List<IRange> keepTogetherRanges = new ArrayList<IRange>();
        //The row 37th and 38th need to keep together. So the pagination results are: the first page if from row 1th to 36th, the second page is from row 37th to 73th.
        keepTogetherRanges.add(worksheet.getRange("37:38"));

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Get the pagination information of the worksheet.
        List<PageInfo> pages = printManager.paginate(worksheet, keepTogetherRanges, null);

        //Save the pages into pdf file.
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
