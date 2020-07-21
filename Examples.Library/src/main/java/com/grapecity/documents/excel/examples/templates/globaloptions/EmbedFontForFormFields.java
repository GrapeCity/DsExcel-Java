package com.grapecity.documents.excel.examples.templates.globaloptions;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.templates.datasource.SalesRecord;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class EmbedFontForFormFields extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_EmbedFontForFormFields.xlsx");
        workbook.open(templateFile);

        //Init template global settings
        workbook.getNames().add("TemplateOptions.EmbedFontForFormFields", "false");

        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_EmbedFontForFormFields.xlsx";
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
        return new String[] { "xlsx/Template_EmbedFontForFormFields.xlsx" };
	}

    @Override
    public boolean getViewInGcPdfViewer() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}