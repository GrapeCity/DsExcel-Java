package com.grapecity.documents.excel.examples.templates.pdfformbuildersamples;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class USTaxForm_fw4 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_fw4-USTaxForm.xlsx");
        workbook.open(templateFile);
        
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_fw4-USTaxForm.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
        return true;
	}
	
	@Override
	public boolean getShowTemplate() {
        return true;
    }

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/Template_fw4-USTaxForm.xlsx" };
	}

    @Override
    public boolean getViewInGcPdfViewer() {
        return true;
    }

	@Override
	public boolean getCanDownload() {
		return false;
	}
}
