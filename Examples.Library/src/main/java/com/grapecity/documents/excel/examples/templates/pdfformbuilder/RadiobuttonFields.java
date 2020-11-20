package com.grapecity.documents.excel.examples.templates.pdfformbuilder;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class RadiobuttonFields extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_RadiobuttonField.xlsx");
        workbook.open(templateFile);

        //#region Init Data
        List<String> choice = new ArrayList<String>();

        choice.add("Yes");
        choice.add("No");
        choice.add("Not sure");
        //#endregion

        //Init template global settings
        workbook.getNames().add("TemplateOptions.KeepLineSize", "true");
        workbook.getNames().add("TemplateOptions.InsertMode", "EntireRowColumn");
        
        //Add data source
        workbook.addDataSource("choice", choice);
        
        //Invoke to process the template
        workbook.processTemplate();
	}
	
	@Override
	public String getTemplateName() {
        return "Template_RadiobuttonField.xlsx";
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
        return new String[] { "xlsx/Template_RadiobuttonField.xlsx" };
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
