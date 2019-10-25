package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Sales extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//Load template file Template_Score.xlsx from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_Sales.xlsx");
		workbook.open(templateFile);

		/// #region Define custom classes
		//		public class SalesInfo {
		//			public String area;
		//			public String salesman;
		//			public String product;
		//			public String productType;
		//			public int sales;
		//		}
		/// #endregion

		///#region Init Data
		List<SalesInfo> salesInfos = new ArrayList<SalesInfo>();

		SalesInfo salesInfo1 = new SalesInfo();
		salesInfo1.area = "NorthChina";
		salesInfo1.salesman = "Hellen";
		salesInfo1.product = "Apple";
		salesInfo1.productType = "Fruit";
		salesInfo1.sales = 120;
		salesInfos.add(salesInfo1);

		SalesInfo salesInfo2 = new SalesInfo();
		salesInfo2.area = "NorthChina";
		salesInfo2.salesman = "Hellen";
		salesInfo2.product = "Banana";
		salesInfo2.productType = "Fruit";
		salesInfo2.sales = 143;
		salesInfos.add(salesInfo2);

		SalesInfo salesInfo3 = new SalesInfo();
		salesInfo3.area = "NorthChina";
		salesInfo3.salesman = "Hellen";
		salesInfo3.product = "Kiwi";
		salesInfo3.productType = "Fruit";
		salesInfo3.sales = 322;
		salesInfos.add(salesInfo3);

		SalesInfo salesInfo4 = new SalesInfo();
		salesInfo4.area = "NorthChina";
		salesInfo4.salesman = "Hellen";
		salesInfo4.product = "Carrots";
		salesInfo4.productType = "Vegetable";
		salesInfo4.sales = 154;
		salesInfos.add(salesInfo4);

		SalesInfo salesInfo5 = new SalesInfo();
		salesInfo5.area = "NorthChina";
		salesInfo5.salesman = "Fancy";
		salesInfo5.product = "Carrots";
		salesInfo5.productType = "Vegetable";
		salesInfo5.sales = 131;
		salesInfos.add(salesInfo5);

		SalesInfo salesInfo6 = new SalesInfo();
		salesInfo6.area = "NorthChina";
		salesInfo6.salesman = "Fancy";
		salesInfo6.product = "Cabbage";
		salesInfo6.productType = "Vegetable";
		salesInfo6.sales = 98;
		salesInfos.add(salesInfo6);

		SalesInfo salesInfo7 = new SalesInfo();
		salesInfo7.area = "NorthChina";
		salesInfo7.salesman = "Fancy";
		salesInfo7.product = "Potato";
		salesInfo7.productType = "Vegetable";
		salesInfo7.sales = 212;
		salesInfos.add(salesInfo7);

		SalesInfo salesInfo8 = new SalesInfo();
		salesInfo8.area = "NorthChina";
		salesInfo8.salesman = "Fancy";
		salesInfo8.product = "Apple";
		salesInfo8.productType = "Fruit";
		salesInfo8.sales = 120;
		salesInfos.add(salesInfo8);

		SalesInfo salesInfo9 = new SalesInfo();
		salesInfo9.area = "NorthChina";
		salesInfo9.salesman = "Ivan";
		salesInfo9.product = "Apple";
		salesInfo9.productType = "Fruit";
		salesInfo9.sales = 164;
		salesInfos.add(salesInfo9);

		SalesInfo salesInfo10 = new SalesInfo();
		salesInfo10.area = "NorthChina";
		salesInfo10.salesman = "Ivan";
		salesInfo10.product = "Kiwi";
		salesInfo10.productType = "Fruit";
		salesInfo10.sales = 213;
		salesInfos.add(salesInfo10);

		SalesInfo salesInfo11 = new SalesInfo();
		salesInfo11.area = "NorthChina";
		salesInfo11.salesman = "Ivan";
		salesInfo11.product = "Potato";
		salesInfo11.productType = "Vegetable";
		salesInfo11.sales = 56;
		salesInfos.add(salesInfo11);

		SalesInfo salesInfo12 = new SalesInfo();
		salesInfo12.area = "NorthChina";
		salesInfo12.salesman = "Ivan";
		salesInfo12.product = "Cabbage";
		salesInfo12.productType = "Vegetable";
		salesInfo12.sales = 265;
		salesInfos.add(salesInfo12);

		SalesInfo salesInfo13 = new SalesInfo();
		salesInfo13.area = "SouthChina";
		salesInfo13.salesman = "Adam";
		salesInfo13.product = "Cabbage";
		salesInfo13.productType = "Vegetable";
		salesInfo13.sales = 112;
		salesInfos.add(salesInfo13);

		SalesInfo salesInfo14 = new SalesInfo();
		salesInfo14.area = "SouthChina";
		salesInfo14.salesman = "Adam";
		salesInfo14.product = "Carrots";
		salesInfo14.productType = "Vegetable";
		salesInfo14.sales = 354;
		salesInfos.add(salesInfo14);

		SalesInfo salesInfo15 = new SalesInfo();
		salesInfo15.area = "SouthChina";
		salesInfo15.salesman = "Adam";
		salesInfo15.product = "Banana";
		salesInfo15.productType = "Fruit";
		salesInfo15.sales = 277;
		salesInfos.add(salesInfo15);

		SalesInfo salesInfo16 = new SalesInfo();
		salesInfo16.area = "SouthChina";
		salesInfo16.salesman = "Adam";
		salesInfo16.product = "Apple";
		salesInfo16.productType = "Fruit";
		salesInfo16.sales = 105;
		salesInfos.add(salesInfo16);

		SalesInfo salesInfo17 = new SalesInfo();
		salesInfo17.area = "SouthChina";
		salesInfo17.salesman = "Bob";
		salesInfo17.product = "Kiwi";
		salesInfo17.productType = "Fruit";
		salesInfo17.sales = 402;
		salesInfos.add(salesInfo17);

		SalesInfo salesInfo18 = new SalesInfo();
		salesInfo18.area = "SouthChina";
		salesInfo18.salesman = "Bob";
		salesInfo18.product = "Banana";
		salesInfo18.productType = "Fruit";
		salesInfo18.sales = 133;
		salesInfos.add(salesInfo1);

		SalesInfo salesInfo19 = new SalesInfo();
		salesInfo19.area = "SouthChina";
		salesInfo19.salesman = "Bob";
		salesInfo19.product = "Cabbage";
		salesInfo19.productType = "Vegetable";
		salesInfo19.sales = 252;
		salesInfos.add(salesInfo19);

		SalesInfo salesInfo20 = new SalesInfo();
		salesInfo20.area = "SouthChina";
		salesInfo20.salesman = "Bob";
		salesInfo20.product = "Potato";
		salesInfo20.productType = "Vegetable";
		salesInfo20.sales = 265;
		salesInfos.add(salesInfo20);
		///#endregion

		//Add data source
		workbook.addDataSource("ds", salesInfos);
		//Invoke to process the template
		workbook.processTemplate();
	}

	@Override
	public boolean getIsNew() {
		return true;
	}

	@Override
	public String getTemplateName() {
		return "Template_Sales.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
		return true;
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/Template_Sales.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
		return new String[] { "SalesInfo" };
	}
}