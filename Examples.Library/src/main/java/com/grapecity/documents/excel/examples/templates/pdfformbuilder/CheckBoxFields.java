package com.grapecity.documents.excel.examples.templates.pdfformbuilder;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class CheckBoxFields extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_CheckBoxField.xlsx");
        workbook.open(templateFile);

        //#region Init Data
        List<String> hobbies = new ArrayList<String>();

        hobbies.add("Cooking");
        hobbies.add("Music");
        hobbies.add("Reading");
        hobbies.add("Sports");
        hobbies.add("Travelling");
        //#endregion

        //Add data source
        workbook.addDataSource("hobbies", hobbies);
        
        //Invoke to process the template
        workbook.processTemplate();
	}
	
	@Override
	public String getTemplateName() {
        return "Template_CheckBoxField.xlsx";
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
        return new String[] { "xlsx/Template_CheckBoxField.xlsx" };
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
