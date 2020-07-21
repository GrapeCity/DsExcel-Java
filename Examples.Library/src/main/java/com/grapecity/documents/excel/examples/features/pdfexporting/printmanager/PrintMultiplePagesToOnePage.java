package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageInfo;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PrintMultiplePagesToOnePage extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/Multiple sheets one page.xlsx"));

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Create a pdf document.
        PDDocument doc = new PDDocument();
        //This page will save datas for multiple pages.
        PDPage page = new PDPage();
        doc.addPage(page);

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();
        
        //Get the pagination information of the workbook.
        List<PageInfo> pages = printManager.paginate(workbook);

        //Divide the multiple pages into 1 rows and 2 columns and printed them on one page.
        printManager.draw(doc, page, pages, 1, 2);

        //Save the modified pages into pdf file.
        try {
        	doc.save(outputStream);
        	doc.close();
        } catch (IOException e) {
        	e.printStackTrace();
        }
        
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
        return "Multiple sheets one page.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/Multiple sheets one page.xlsx" };
	}

}
