package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PaginationOrientation;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;

public class GetPaginationInfo extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Medical office start-up expenses 1.xlsx");
        workbook.open(fileStream);
        
        IWorksheet worksheet = workbook.getWorksheets().get(1);

        PrintManager printManager = new PrintManager();

        // The columnIndexs is [4, 12, 20], this means that the horizontal direction is split after the column 5th, 13th, and 21th.  
        List<Integer> columnIndexs = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Horizontal);
        // The rowIndexs is [42, 61], this means that the vertical direction is split after the row 43th and 62th.
        List<Integer> rowIndexs = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Vertical);

        // Save the pages into pdf file.
        List<PageInfo> pages = printManager.paginate(worksheet);
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
	public String[] getResources() {
        return new String[] { "xlsx/Medical office start-up expenses 1.xlsx"};
	}
}