package com.grapecity.documents.excel.examples.features.worksheets;

import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CopyWorksheet extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file AgingReport.xlsx from resource
        InputStream fileStream = this.getResourceStream("xlsx/AgingReport.xlsx");
        workbook.open(fileStream);

        //Copy the active sheet to the end of current workbook
        IWorksheet copy_worksheet = workbook.getActiveSheet().copy();
        copy_worksheet.setName("Copy of " + workbook.getActiveSheet().getName());
	}

	@Override
	public String getTemplateName() {
		return "AgingReport.xlsx";
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/AgingReport.xlsx" };
	}
}