package com.grapecity.documents.excel.examples.templates.pdfformbuildersamples;

import java.io.InputStream;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SendaiTaxForm extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_SendaiTaxForm.xlsx");
        workbook.open(templateFile);
        
        //Invoke to process the template
        workbook.processTemplate();
	}
	
	@Override
	public String getTemplateName() {
        return "Template_SendaiTaxForm.xlsx";
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
        return new String[] { "xlsx/Template_SendaiTaxForm.xlsx" };
	}

	@Override
	public boolean getIsNew() {
		return true;
	}

    @Override
    public boolean getViewInGcPdfViewer() {
        return true;
    }
}
