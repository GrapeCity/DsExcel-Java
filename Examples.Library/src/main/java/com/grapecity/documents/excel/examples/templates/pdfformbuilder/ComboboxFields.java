package com.grapecity.documents.excel.examples.templates.pdfformbuilder;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class ComboboxFields extends ExampleBase {
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_ComboBoxField.xlsx");
        workbook.open(templateFile);

        //#region Init Data
        List<String> names = new ArrayList<String>();

        names.add("Emma");
        names.add("Ava");
        names.add("William");
        names.add("Liam");
        names.add("Noah");
        //#endregion

        //Init template global settings
        workbook.getNames().add("TemplateOptions.KeepLineSize", "true");
        
        //Add data source
        workbook.addDataSource("name", names);
        
        //Invoke to process the template
        workbook.processTemplate();
	}
	
	@Override
	public String getTemplateName() {
        return "Template_ComboBoxField.xlsx";
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
        return new String[] { "xlsx/Template_ComboBoxField.xlsx" };
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
