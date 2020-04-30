package com.grapecity.documents.excel.examples.templates;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.templates.datasource.SalesRecord;

public class TemplateWorkWithChart extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_WorkWithChart.xlsx");
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

        SalesRecord record2 = new SalesRecord();
        record2.area = "North America";
        record2.city = "New York";
        record2.category = "Consumer Electronics";
        record2.name = "Bose 785593-0050";
        record2.revenue = 92800;
        datasource.add(record2);

        SalesRecord record3 = new SalesRecord();
        record3.area = "South America";
        record3.city = "Santiago";
        record3.category = "Consumer Electronics";
        record3.name = "Bose 785593-0050";
        record3.revenue = 19550;
        datasource.add(record3);

        SalesRecord record4 = new SalesRecord();
        record4.area = "North America";
        record4.city = "Chicago";
        record4.category = "Consumer Electronics";
        record4.name = "Canon EOS 1500D";
        record4.revenue = 98650;
        datasource.add(record4);

        SalesRecord record5 = new SalesRecord();
        record5.area = "North America";
        record5.city = "Minnesota";
        record5.category = "Consumer Electronics";
        record5.name = "Canon EOS 1500D";
        record5.revenue = 89110;
        datasource.add(record5);

        SalesRecord record6 = new SalesRecord();
        record6.area = "South America";
        record6.city = "Santiago";
        record6.category = "Consumer Electronics";
        record6.name = "Canon EOS 1500D";
        record6.revenue = 459000;
        datasource.add(record6);

        SalesRecord record7 = new SalesRecord();
        record7.area = "North America";
        record7.city = "Chicago";
        record7.category = "Consumer Electronics";
        record7.name = "Haier 394L 4Star";
        record7.revenue = 367050;
        datasource.add(record7);

        SalesRecord record8 = new SalesRecord();
        record8.area = "South America";
        record8.city = "Quito";
        record8.category = "Consumer Electronics";
        record8.name = "Haier 394L 4Star";
        record8.revenue = 729100;
        datasource.add(record8);

        SalesRecord record9 = new SalesRecord();
        record9.area = "South America";
        record9.city = "Santiago";
        record9.category = "Consumer Electronics";
        record9.name = "Haier 394L 4Star";
        record9.revenue = 578900;
        datasource.add(record9);

        SalesRecord record10 = new SalesRecord();
        record10.area = "North America";
        record10.city = "Fremont";
        record10.category = "Consumer Electronics";
        record10.name = "IFB 6.5 Kg FullyAuto";
        record10.revenue = 904930;
        datasource.add(record10);

        SalesRecord record11 = new SalesRecord();
        record11.area = "South America";
        record11.city = "Buenos Aires";
        record11.category = "Consumer Electronics";
        record11.name = "IFB 6.5 Kg FullyAuto";
        record11.revenue = 673800;
        datasource.add(record11);

        SalesRecord record12 = new SalesRecord();
        record12.area = "South America";
        record12.city = "Medillin";
        record12.category = "Consumer Electronics";
        record12.name = "IFB 6.5 Kg FullyAuto";
        record12.revenue = 82910;
        datasource.add(record12);

        SalesRecord record13 = new SalesRecord();
        record13.area = "North America";
        record13.city = "Chicago";
        record13.category = "Consumer Electronics";
        record13.name = "Mi LED 40inch";
        record13.revenue = 550010;
        datasource.add(record13);

        SalesRecord record14 = new SalesRecord();
        record14.area = "North America";
        record14.city = "Minnesota";
        record14.category = "Consumer Electronics";
        record14.name = "Mi LED 40inch";
        record14.revenue = 1784702;
        datasource.add(record14);

        SalesRecord record15 = new SalesRecord();
        record15.area = "South America";
        record15.city = "Santiago";
        record15.category = "Consumer Electronics";
        record15.name = "Mi LED 40inch";
        record15.revenue = 102905;
        datasource.add(record15);

        SalesRecord record16 = new SalesRecord();
        record16.area = "North America";
        record16.city = "Chicago";
        record16.category = "Consumer Electronics";
        record16.name = "Sennheiser HD 4.40-BT";
        record16.revenue = 178100;
        datasource.add(record16);

        SalesRecord record17 = new SalesRecord();
        record17.area = "South America";
        record17.city = "Quito";
        record17.category = "Consumer Electronics";
        record17.name = "Sennheiser HD 4.40-BT";
        record17.revenue = 234459;
        datasource.add(record17);

        SalesRecord record18 = new SalesRecord();
        record18.area = "North America";
        record18.city = "Minnesota";
        record18.category = "Mobile";
        record18.name = "Iphone XR";
        record18.revenue = 1734621;
        datasource.add(record18);

        SalesRecord record19 = new SalesRecord();
        record19.area = "South America";
        record19.city = "Santiago";
        record19.category = "Mobile";
        record19.name = "Iphone XR";
        record19.revenue = 109300;
        datasource.add(record19);

        SalesRecord record20 = new SalesRecord();
        record20.area = "North America";
        record20.city = "Chicago";
        record20.category = "Mobile";
        record20.name = "OnePlus 7Pro";
        record20.revenue = 499100;
        datasource.add(record20);

        SalesRecord record21 = new SalesRecord();
        record21.area = "South America";
        record21.city = "Quito";
        record21.category = "Mobile";
        record21.name = "OnePlus 7Pro";
        record21.revenue = 215000;
        datasource.add(record21);

        SalesRecord record22 = new SalesRecord();
        record22.area = "North America";
        record22.city = "Minnesota";
        record22.category = "Mobile";
        record22.name = "Redmi 7";
        record22.revenue = 81650;
        datasource.add(record22);

        SalesRecord record23 = new SalesRecord();
        record23.area = "South America";
        record23.city = "Quito";
        record23.category = "Mobile";
        record23.name = "Redmi 7";
        record23.revenue = 276390;
        datasource.add(record23);

        SalesRecord record24 = new SalesRecord();
        record24.area = "North America";
        record24.city = "Minnesota";
        record24.category = "Mobile";
        record24.name = "Samsung S9";
        record24.revenue = 896250;
        datasource.add(record24);

        SalesRecord record25 = new SalesRecord();
        record25.area = "South America";
        record25.city = "Buenos Aires";
        record25.category = "Mobile";
        record25.name = "Samsung S9";
        record25.revenue = 896250;
        datasource.add(record25);

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
        return "Template_WorkWithChart.xlsx";
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
        return new String[] { "xlsx/Template_WorkWithChart.xlsx" };
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