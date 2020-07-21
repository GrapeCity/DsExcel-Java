package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class TablixReport extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_TablixReport.xlsx");
        workbook.open(templateFile);

        /// #region Define custom classes
        //        public class TablixInfo {
        //        	public int orderID;
        //        	public String product;
        //        	public double sales;
        //        	public String productType;
        //        	public String year;
        //        	public String season;
        //        }
        /// #endregion

        ///#region Init Data
        List<TablixInfo> tablixInfos = new ArrayList<TablixInfo>();

        TablixInfo tablixInfo1 = new TablixInfo();
        tablixInfo1.orderID = 1;
        tablixInfo1.product = "Röd Kaviar";
        tablixInfo1.sales = 300;
        tablixInfo1.productType = "Seafood";
        tablixInfo1.year = "2017";
        tablixInfo1.season = "Q3";
        tablixInfos.add(tablixInfo1);

        TablixInfo tablixInfo2 = new TablixInfo();
        tablixInfo2.orderID = 2;
        tablixInfo2.product = "Spegesild";
        tablixInfo2.sales = 144;
        tablixInfo2.productType = "Seafood";
        tablixInfo2.year = "2017";
        tablixInfo2.season = "Q3";
        tablixInfos.add(tablixInfo2);

        TablixInfo tablixInfo3 = new TablixInfo();
        tablixInfo3.orderID = 3;
        tablixInfo3.product = "Carnarvon Tigers";
        tablixInfo3.sales = 600;
        tablixInfo3.productType = "Seafood";
        tablixInfo3.year = "2017";
        tablixInfo3.season = "Q3";
        tablixInfos.add(tablixInfo3);

        TablixInfo tablixInfo4 = new TablixInfo();
        tablixInfo4.orderID = 4;
        tablixInfo4.product = "Spegesild";
        tablixInfo4.sales = 288;
        tablixInfo4.productType = "Seafood";
        tablixInfo4.year = "2017";
        tablixInfo4.season = "Q4";
        tablixInfos.add(tablixInfo4);

        TablixInfo tablixInfo5 = new TablixInfo();
        tablixInfo5.orderID = 5;
        tablixInfo5.product = "Carnarvon Tigers";
        tablixInfo5.sales = 4250;
        tablixInfo5.productType = "Seafood";
        tablixInfo5.year = "2017";
        tablixInfo5.season = "Q4";
        tablixInfos.add(tablixInfo5);

        TablixInfo tablixInfo6 = new TablixInfo();
        tablixInfo6.orderID = 6;
        tablixInfo6.product = "Escargots de Bourgogne";
        tablixInfo6.sales = 636;
        tablixInfo6.productType = "Seafood";
        tablixInfo6.year = "2017";
        tablixInfo6.season = "Q4";
        tablixInfos.add(tablixInfo6);

        TablixInfo tablixInfo7 = new TablixInfo();
        tablixInfo7.orderID = 7;
        tablixInfo7.product = "Röd Kaviar";
        tablixInfo7.sales = 240;
        tablixInfo7.productType = "Seafood";
        tablixInfo7.year = "2018";
        tablixInfo7.season = "Q1";
        tablixInfos.add(tablixInfo7);

        TablixInfo tablixInfo8 = new TablixInfo();
        tablixInfo8.orderID = 8;
        tablixInfo8.product = "Carnarvon Tigers";
        tablixInfo8.sales = 450;
        tablixInfo8.productType = "Seafood";
        tablixInfo8.year = "2018";
        tablixInfo8.season = "Q1";
        tablixInfos.add(tablixInfo8);

        TablixInfo tablixInfo9 = new TablixInfo();
        tablixInfo9.orderID = 9;
        tablixInfo9.product = "Röd Kaviar";
        tablixInfo9.sales = 735;
        tablixInfo9.productType = "Seafood";
        tablixInfo9.year = "2018";
        tablixInfo9.season = "Q2";
        tablixInfos.add(tablixInfo9);

        TablixInfo tablixInfo10 = new TablixInfo();
        tablixInfo10.orderID = 10;
        tablixInfo10.product = "Røgede sild";
        tablixInfo10.sales = 1377;
        tablixInfo10.productType = "Seafood";
        tablixInfo10.year = "2018";
        tablixInfo10.season = "Q2";
        tablixInfos.add(tablixInfo10);

        TablixInfo tablixInfo11 = new TablixInfo();
        tablixInfo11.orderID = 11;
        tablixInfo11.product = "Röd Kaviar";
        tablixInfo11.sales = 1020;
        tablixInfo11.productType = "Seafood";
        tablixInfo11.year = "2018";
        tablixInfo11.season = "Q3";
        tablixInfos.add(tablixInfo11);

        TablixInfo tablixInfo12 = new TablixInfo();
        tablixInfo12.orderID = 12;
        tablixInfo12.product = "Røgede sild";
        tablixInfo12.sales = 190;
        tablixInfo12.productType = "Seafood";
        tablixInfo12.year = "2018";
        tablixInfo12.season = "Q3";
        tablixInfos.add(tablixInfo12);

        TablixInfo tablixInfo13 = new TablixInfo();
        tablixInfo13.orderID = 13;
        tablixInfo13.product = "Röd Kaviar";
        tablixInfo13.sales = 1725;
        tablixInfo13.productType = "Seafood";
        tablixInfo13.year = "2018";
        tablixInfo13.season = "Q4";
        tablixInfos.add(tablixInfo13);

        TablixInfo tablixInfo14 = new TablixInfo();
        tablixInfo14.orderID = 14;
        tablixInfo14.product = "Carnarvon Tigers";
        tablixInfo14.sales = 3562;
        tablixInfo14.productType = "Seafood";
        tablixInfo14.year = "2018";
        tablixInfo14.season = "Q4";
        tablixInfos.add(tablixInfo14);

        TablixInfo tablixInfo15 = new TablixInfo();
        tablixInfo15.orderID = 15;
        tablixInfo15.product = "Sir Rodney's Marmalade";
        tablixInfo15.sales = 4276;
        tablixInfo15.productType = "Confections";
        tablixInfo15.year = "2017";
        tablixInfo15.season = "Q3";
        tablixInfos.add(tablixInfo15);

        TablixInfo tablixInfo16 = new TablixInfo();
        tablixInfo16.orderID = 16;
        tablixInfo16.product = "Maxilaku";
        tablixInfo16.sales = 880;
        tablixInfo16.productType = "Confections";
        tablixInfo16.year = "2017";
        tablixInfo16.season = "Q3";
        tablixInfos.add(tablixInfo16);

        TablixInfo tablixInfo17 = new TablixInfo();
        tablixInfo17.orderID = 17;
        tablixInfo17.product = "Maxilaku";
        tablixInfo17.sales = 1040;
        tablixInfo17.productType = "Confections";
        tablixInfo17.year = "2017";
        tablixInfo17.season = "Q4";
        tablixInfos.add(tablixInfo17);

        TablixInfo tablixInfo18 = new TablixInfo();
        tablixInfo18.orderID = 18;
        tablixInfo18.product = "NuNuCa Nuß-Nougat-Creme";
        tablixInfo18.sales = 716.8;
        tablixInfo18.productType = "Confections";
        tablixInfo18.year = "2017";
        tablixInfo18.season = "Q4";
        tablixInfos.add(tablixInfo18);

        TablixInfo tablixInfo19 = new TablixInfo();
        tablixInfo19.orderID = 1;
        tablixInfo19.product = "Sir Rodney's Marmalade";
        tablixInfo19.sales = 2592;
        tablixInfo19.productType = "Confections";
        tablixInfo19.year = "2018";
        tablixInfo19.season = "Q1";
        tablixInfos.add(tablixInfo19);

        TablixInfo tablixInfo20 = new TablixInfo();
        tablixInfo20.orderID = 20;
        tablixInfo20.product = "Maxilaku";
        tablixInfo20.sales = 1296;
        tablixInfo20.productType = "Confections";
        tablixInfo20.year = "2018";
        tablixInfo20.season = "Q1";
        tablixInfos.add(tablixInfo20);

        TablixInfo tablixInfo21 = new TablixInfo();
        tablixInfo21.orderID = 21;
        tablixInfo21.product = "Pavlova";
        tablixInfo21.sales = 1473.4;
        tablixInfo21.productType = "Confections";
        tablixInfo21.year = "2018";
        tablixInfo21.season = "Q1";
        tablixInfos.add(tablixInfo21);

        TablixInfo tablixInfo22 = new TablixInfo();
        tablixInfo22.orderID = 22;
        tablixInfo22.product = "Sir Rodney's Marmalade";
        tablixInfo22.sales = 4374;
        tablixInfo22.productType = "Confections";
        tablixInfo22.year = "2018";
        tablixInfo22.season = "Q2";
        tablixInfos.add(tablixInfo22);

        TablixInfo tablixInfo23 = new TablixInfo();
        tablixInfo23.orderID = 23;
        tablixInfo23.product = "Maxilaku";
        tablixInfo23.sales = 1004;
        tablixInfo23.productType = "Confections";
        tablixInfo23.year = "2018";
        tablixInfo23.season = "Q2";
        tablixInfos.add(tablixInfo1);

        TablixInfo tablixInfo24 = new TablixInfo();
        tablixInfo24.orderID = 24;
        tablixInfo24.product = "Pavlova";
        tablixInfo24.sales = 3075;
        tablixInfo24.productType = "Confections";
        tablixInfo24.year = "2018";
        tablixInfo24.season = "Q2";
        tablixInfos.add(tablixInfo24);

        TablixInfo tablixInfo25 = new TablixInfo();
        tablixInfo25.orderID = 25;
        tablixInfo25.product = "Sir Rodney's Marmalade";
        tablixInfo25.sales = 1071;
        tablixInfo25.productType = "Confections";
        tablixInfo25.year = "2018";
        tablixInfo25.season = "Q3";
        tablixInfos.add(tablixInfo25);

        TablixInfo tablixInfo26 = new TablixInfo();
        tablixInfo26.orderID = 1;
        tablixInfo26.product = "Maxilaku";
        tablixInfo26.sales = 860;
        tablixInfo26.productType = "Confections";
        tablixInfo26.year = "2018";
        tablixInfo26.season = "Q3";
        tablixInfos.add(tablixInfo26);

        TablixInfo tablixInfo27 = new TablixInfo();
        tablixInfo27.orderID = 27;
        tablixInfo27.product = "Pavlova";
        tablixInfo27.sales = 732;
        tablixInfo27.productType = "Confections";
        tablixInfo27.year = "2018";
        tablixInfo27.season = "Q3";
        tablixInfos.add(tablixInfo27);

        TablixInfo tablixInfo28 = new TablixInfo();
        tablixInfo28.orderID = 28;
        tablixInfo28.product = "Sir Rodney's Marmalade";
        tablixInfo28.sales = 1071;
        tablixInfo28.productType = "Confections";
        tablixInfo28.year = "2018";
        tablixInfo28.season = "Q4";
        tablixInfos.add(tablixInfo28);

        TablixInfo tablixInfo29 = new TablixInfo();
        tablixInfo29.orderID = 29;
        tablixInfo29.product = "Pavlova";
        tablixInfo29.sales = 2634;
        tablixInfo29.productType = "Confections";
        tablixInfo29.year = "2018";
        tablixInfo29.season = "Q4";
        tablixInfos.add(tablixInfo29);

        TablixInfo tablixInfo30 = new TablixInfo();
        tablixInfo30.orderID = 30;
        tablixInfo30.product = "Sir Rodney's Scones";
        tablixInfo30.sales = 1790;
        tablixInfo30.productType = "Confections";
        tablixInfo30.year = "2018";
        tablixInfo30.season = "Q4";
        tablixInfos.add(tablixInfo30);
        ///#endregion

        //Add data source
        workbook.addDataSource("ds", tablixInfos);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_TablixReport.xlsx";
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
        return new String[] { "xlsx/Template_TablixReport.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "TablixInfo" };
	}
}