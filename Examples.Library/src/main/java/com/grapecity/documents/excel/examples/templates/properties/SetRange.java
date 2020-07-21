package com.grapecity.documents.excel.examples.templates.properties;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.templates.datasource.SalesRecord;

public class SetRange extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file Template_Score.xlsx from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_SetRange.xlsx");
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
        SalesRecord record1 = new SalesRecord();
        record1.area = "North America";
        record1.city = "Chicago";
        record1.category = "Consumer Electronics";
        record1.name = "Bose 785593-0050";
        record1.revenue = 92800;
        datasource.add(record1);

        SalesRecord record5 = new SalesRecord();
        record5.area = "North America";
        record5.city = "Minnesota";
        record5.category = "Consumer Electronics";
        record5.name = "Canon EOS 1500D";
        record5.revenue = 89110;
        datasource.add(record5);

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

        SalesRecord record16 = new SalesRecord();
        record16.area = "North America";
        record16.city = "Chicago";
        record16.category = "Consumer Electronics";
        record16.name = "Sennheiser HD 4.40-BT";
        record16.revenue = 178100;
        datasource.add(record16);

        SalesRecord record18 = new SalesRecord();
        record18.area = "North America";
        record18.city = "Minnesota";
        record18.category = "Mobile";
        record18.name = "Iphone XR";
        record18.revenue = 1734621;
        datasource.add(record18);

        SalesRecord record21 = new SalesRecord();
        record21.area = "South America";
        record21.city = "Quito";
        record21.category = "Mobile";
        record21.name = "OnePlus 7Pro";
        record21.revenue = 215000;
        datasource.add(record21);

        SalesRecord record23 = new SalesRecord();
        record23.area = "South America";
        record23.city = "Quito";
        record23.category = "Mobile";
        record23.name = "Redmi 7";
        record23.revenue = 276390;
        datasource.add(record23);

        SalesRecord record25 = new SalesRecord();
        record25.area = "South America";
        record25.city = "Buenos Aires";
        record25.category = "Mobile";
        record25.name = "Samsung S9";
        record25.revenue = 896250;
        datasource.add(record25);
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
        return "Template_SetRange.xlsx";
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
        return new String[] { "xlsx/Template_SetRange.xlsx" };
	}

	@Override
	public String[] getRefs() {
        return new String[] { "SalesData", "SalesRecord" };
	}
}