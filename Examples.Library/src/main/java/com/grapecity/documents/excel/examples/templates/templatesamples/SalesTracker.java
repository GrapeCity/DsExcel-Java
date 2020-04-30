package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.GcMockResultSet;

public class SalesTracker extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_SalesTracker.xlsx");
        workbook.open(templateFile);

        //Here in the demo, we use a mock class to generate instance of java.sql.ResultSet.
        //User who use template in product, must get instance of java.sql.ResultSet from the 
        //related database connection.
        java.sql.ResultSet datasource = new GcMockResultSet(this.getResourceStream("data/sales tracker.csv"));

        //Init template global settings
        workbook.getNames().add("TemplateOptions.KeepLineSize", "true");
        
        //Add data source
        workbook.addDataSource("ds", datasource);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_SalesTracker.xlsx";
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
        return new String[] { "xlsx/Template_SalesTracker.xlsx", "data/sales tracker.csv" };
	}
	
	@Override
	public boolean getIsNew() {
        return true;
	}
}