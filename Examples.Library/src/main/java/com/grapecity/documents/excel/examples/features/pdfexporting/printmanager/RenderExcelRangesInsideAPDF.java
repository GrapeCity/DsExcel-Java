package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Rectangle;
import com.grapecity.documents.excel.Size;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class RenderExcelRangesInsideAPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/FinancialReport.xlsx"));
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        //Create a pdf document.
        PDDocument doc = null;
        try {
        	doc = PDDocument.load(this.getResourceStream("Acme-Financial Report 2018.pdf"));
        } catch (IOException e1) {
        	e1.printStackTrace();
        }
   
        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Draw the contents of the sheet3 to the fourth page. 
        IRange printArea1 = workbook.getWorksheets().get(2).getRange("A3:C24");
        Size size1 = printManager.getSize(printArea1);
        printManager.draw(doc, doc.getPage(3), new Rectangle(306, 215, size1.getWidth(), size1.getHeight()), printArea1);

        //Draw the contents of the sheet1 to the fifth page. 
        IRange printArea2 = workbook.getWorksheets().get(0).getRange("A4:E29");
        Size size2 = printManager.getSize(printArea2);
        printManager.draw(doc, doc.getPage(4), new Rectangle(71, 250, size2.getWidth(), size2.getHeight()), printArea2);

        //Draw the contents of the sheet2 to the sixth page. 
        IRange printArea3 = workbook.getWorksheets().get(1).getRange("A2:E28");
        Size size3 = printManager.getSize(printArea3);
        printManager.draw(doc, doc.getPage(5), new Rectangle(71, 230, 783, size3.getHeight()), printArea3);    

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
        return "FinancialReport.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] {  "xlsx/FinancialReport.xlsx", "Acme-Financial Report 2018.pdf" };
	}

}
