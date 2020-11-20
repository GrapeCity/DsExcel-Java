package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.InputStream;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertExcelToPDF extends ExampleBase {
	@Override
    public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Employee absence schedule.xlsx");
        workbook.open(fileStream);
	}

	@Override
	public boolean getSavePdf() {
        return true;
	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

	@Override
	public String getTemplateName() {
        return "Employee absence schedule.xlsx";
	}

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/Employee absence schedule.xlsx" };
	}
}