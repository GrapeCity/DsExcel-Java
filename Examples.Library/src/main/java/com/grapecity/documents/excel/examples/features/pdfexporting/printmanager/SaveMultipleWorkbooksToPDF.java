package com.grapecity.documents.excel.examples.features.pdfexporting.printmanager;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.IWorkbook;
import com.grapecity.documents.excel.PrintManager;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveMultipleWorkbooksToPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        workbook.open(this.getResourceStream("xlsx/Any year calendar1.xlsx"));

        Workbook workbook2 = new Workbook();
        workbook2.open(this.getResourceStream("xlsx/Any year calendar (Ion theme)1.xlsx"));

        //Create a PrintManager.
        PrintManager printManager = new PrintManager();

        //Save the workbook1 and workbook2 into one pdf file.
        List<IWorkbook> workbooks = new ArrayList<IWorkbook>();
        workbooks.add(workbook);
        workbooks.add(workbook2);
        printManager.saveWorkbooksToPDF(outputStream, workbooks);
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
		return "Any year calendar1.xlsx";
	}

	@Override
	public String[] getResources() {
		return new String[] {  "xlsx/Any year calendar1.xlsx" , "xlsx/Any year calendar (Ion theme)1.xlsx"  };
	}

}
