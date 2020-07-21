package com.grapecity.documents.excel.examples.templates.properties.fillmode;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.templates.datasource.SalesRecord;

public class SetFillModeOverwrite extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_SetFillModeOverwrite.xlsx");
        workbook.open(templateFile);

        //#region Define custom classes
        //public class SalesData
        //{
        //    public List<SalesRecord> sales;
        //}

        //public class SalesRecord
        //{
        //    public String area;
        //    public String city;
        //    public String category;
        //    public String name;
        //    public double revenue;
        //}
        //#endregion

        List<SalesRecord> datasource = new ArrayList<SalesRecord>();

        //#region Init Data
        SalesRecord record7 = new SalesRecord();
        record7.area = "North America";
        record7.city = "Chicago";
        record7.category = "Consumer Electronics";
        record7.name = "Haier 394L 4Star";
        record7.revenue = 367050;
        datasource.add(record7);

        SalesRecord record9 = new SalesRecord();
        record9.area = "South America";
        record9.city = "Santiago";
        record9.category = "Consumer Electronics";
        record9.name = "Haier 394L 4Star";
        record9.revenue = 578900;
        datasource.add(record9);

        SalesRecord record11 = new SalesRecord();
        record11.area = "South America";
        record11.city = "Buenos Aires";
        record11.category = "Consumer Electronics";
        record11.name = "IFB 6.5 Kg FullyAuto";
        record11.revenue = 673800;
        datasource.add(record11);

        SalesRecord record14 = new SalesRecord();
        record14.area = "North America";
        record14.city = "Minnesota";
        record14.category = "Consumer Electronics";
        record14.name = "Mi LED 40inch";
        record14.revenue = 1784702;
        datasource.add(record14);

        SalesRecord record19 = new SalesRecord();
        record19.area = "South America";
        record19.city = "Santiago";
        record19.category = "Mobile";
        record19.name = "Iphone XR";
        record19.revenue = 109300;
        datasource.add(record19);

        SalesRecord record22 = new SalesRecord();
        record22.area = "North America";
        record22.city = "Minnesota";
        record22.category = "Mobile";
        record22.name = "Redmi 7";
        record22.revenue = 81650;
        datasource.add(record22);

        SalesRecord record24 = new SalesRecord();
        record24.area = "North America";
        record24.city = "Minnesota";
        record24.category = "Mobile";
        record24.name = "Samsung S9";
        record24.revenue = 896250;
        datasource.add(record24);

        SalesRecord record26 = new SalesRecord();
        record26.area = "South America";
        record26.city = "Quito";
        record26.category = "Mobile";
        record26.name = "Samsung S9";
        record26.revenue = 716520;
        datasource.add(record26);
        //#endregion

        //Init template global settings
        workbook.getNames().add("TemplateOptions.KeepLineSize", "true");

        //Add data source
        workbook.addDataSource("ds", datasource);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_SetFillModeOverwrite.xlsx";
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
        return new String[] { "xlsx/Template_SetFillModeOverwrite.xlsx" };
	}

	@Override
	public String[] getRefs() {
        return new String[] { "SalesData", "SalesRecord" };
	}

	@Override
	public boolean getIsNew() {
        return true;
	}
}