package com.grapecity.documents.excel.examples.templates.datasource;

import java.io.InputStream;
import java.util.ArrayList;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CustomObject extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//Load template file Template_SalesDataGroup.xlsx from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_SalesDataGroup.xlsx");
		workbook.open(templateFile);

		/// #region Define custom classes
		//  public class SalesData
		//  {
		// 		public ArrayList<SalesRecord> records;
		//  }

		//  public class SalesRecord
		//  {
		// 		public String area;
		// 		public String salesman;
		// 		public String product;
		// 		public String productType;
		// 		public int sales;
		//  }
		/// #endregion

		SalesData datasource = new SalesData();
		datasource.records = new ArrayList<SalesRecord>();

		///#region Init Data
		SalesRecord record1 = new SalesRecord();
		record1.area = "NorthChina";
		record1.salesman = "Hellen";
		record1.product = "Apple";
		record1.productType = "Fruit";
		record1.sales = 120;
		datasource.records.add(record1);

		SalesRecord record2 = new SalesRecord();
		record2.area = "NorthChina";
		record2.salesman = "Hellen";
		record2.product = "Banana";
		record2.productType = "Fruit";
		record2.sales = 143;
		datasource.records.add(record2);

		SalesRecord record3 = new SalesRecord();
		record3.area = "NorthChina";
		record3.salesman = "Hellen";
		record3.product = "Kiwi";
		record3.productType = "Fruit";
		record3.sales = 322;
		datasource.records.add(record3);

		SalesRecord record4 = new SalesRecord();
		record4.area = "NorthChina";
		record4.salesman = "Hellen";
		record4.product = "Carrots";
		record4.productType = "Vegetable";
		record4.sales = 154;
		datasource.records.add(record4);

		SalesRecord record5 = new SalesRecord();
		record5.area = "NorthChina";
		record5.salesman = "Fancy";
		record5.product = "Carrots";
		record5.productType = "Vegetable";
		record5.sales = 131;
		datasource.records.add(record5);

		SalesRecord record6 = new SalesRecord();
		record6.area = "NorthChina";
		record6.salesman = "Fancy";
		record6.product = "Cabbage";
		record6.productType = "Vegetable";
		record6.sales = 98;
		datasource.records.add(record6);

		SalesRecord record7 = new SalesRecord();
		record7.area = "NorthChina";
		record7.salesman = "Fancy";
		record7.product = "Potato";
		record7.productType = "Vegetable";
		record7.sales = 212;
		datasource.records.add(record7);

		SalesRecord record8 = new SalesRecord();
		record8.area = "NorthChina";
		record8.salesman = "Fancy";
		record8.product = "Apple";
		record8.productType = "Fruit";
		record8.sales = 102;
		datasource.records.add(record8);

		SalesRecord record9 = new SalesRecord();
		record9.area = "NorthChina";
		record9.salesman = "Ivan";
		record9.product = "Apple";
		record9.productType = "Fruit";
		record9.sales = 164;
		datasource.records.add(record9);

		SalesRecord record10 = new SalesRecord();
		record10.area = "NorthChina";
		record10.salesman = "Ivan";
		record10.product = "Kiwi";
		record10.productType = "Fruit";
		record10.sales = 213;
		datasource.records.add(record10);

		SalesRecord record11 = new SalesRecord();
		record11.area = "NorthChina";
		record11.salesman = "Ivan";
		record11.product = "Potato";
		record11.productType = "Vegetable";
		record11.sales = 56;
		datasource.records.add(record11);

		SalesRecord record12 = new SalesRecord();
		record12.area = "NorthChina";
		record12.salesman = "Ivan";
		record12.product = "Cabbage";
		record12.productType = "Vegetable";
		record12.sales = 265;
		datasource.records.add(record12);

		SalesRecord record13 = new SalesRecord();
		record13.area = "SouthChina";
		record13.salesman = "Adam";
		record13.product = "Cabbage";
		record13.productType = "Vegetable";
		record13.sales = 112;
		datasource.records.add(record13);

		SalesRecord record14 = new SalesRecord();
		record14.area = "SouthChina";
		record14.salesman = "Adam";
		record14.product = "Carrots";
		record14.productType = "Vegetable";
		record14.sales = 354;
		datasource.records.add(record14);

		SalesRecord record15 = new SalesRecord();
		record15.area = "SouthChina";
		record15.salesman = "Adam";
		record15.product = "Banana";
		record15.productType = "Fruit";
		record15.sales = 277;
		datasource.records.add(record15);

		SalesRecord record16 = new SalesRecord();
		record16.area = "SouthChina";
		record16.salesman = "Adam";
		record16.product = "Apple";
		record16.productType = "Fruit";
		record16.sales = 105;
		datasource.records.add(record16);

		SalesRecord record17 = new SalesRecord();
		record17.area = "SouthChina";
		record17.salesman = "Bob";
		record17.product = "Banana";
		record17.productType = "Fruit";
		record17.sales = 133;
		datasource.records.add(record17);

		SalesRecord record18 = new SalesRecord();
		record18.area = "SouthChina";
		record18.salesman = "Bob";
		record18.product = "Cabbage";
		record18.productType = "Vegetable";
		record18.sales = 252;
		datasource.records.add(record18);

		SalesRecord record19 = new SalesRecord();
		record19.area = "SouthChina";
		record19.salesman = "Bob";
		record19.product = "Potato";
		record19.productType = "Vegetable";
		record19.sales = 265;
		datasource.records.add(record19);

		SalesRecord record20 = new SalesRecord();
		record20.area = "SouthChina";
		record20.salesman = "Bob";
		record20.product = "Kiwi";
		record20.productType = "Fruit";
		record20.sales = 402;
		datasource.records.add(record20);
		///#endregion

		//Add data source
		workbook.addDataSource("ds", datasource);
		//Invoke to process the template
		workbook.processTemplate();
	}

	@Override
	public boolean getIsNew() {
		return true;
	}

	@Override
	public String getTemplateName() {
		return "Template_SalesDataGroup.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
		return true;
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/Template_SalesDataGroup.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
		return new String[] { "SalesData", "SalesRecord" };
	}
}