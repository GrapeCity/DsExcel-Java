package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.GregorianCalendar;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PurchaseOrder extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_PurchaseOrder.xlsx");
        workbook.open(templateFile);

        /// #region Define custom classes

        //        public class PurchaseOrderInfo {
        //        	public int s_no;
        //        	public String itemnumber;
        //        	public String itemdescription;
        //        	public int quantity;
        //        	public String um;
        //        	public int price;
        //        }

        //        public class PurchaseOrderBasicInfo {
        //        	public String iD;
        //        	public Date orderDate = new Date(0);
        //        	public String creditTerms;
        //        	public String pONumber;
        //        	public String ref;
        //        	public String deliverToCompany;
        //        	public String deliverToAddress;
        //        	public String postalCode;
        //        	public String country;
        //
        //        }
        /// #endregion

        ///#region Init Data
        List<PurchaseOrderInfo> purchaseOrderInfoList = new ArrayList<PurchaseOrderInfo>();

        PurchaseOrderInfo purchaseOrderInfo1 = new PurchaseOrderInfo();
        purchaseOrderInfo1.s_no = 1;
        purchaseOrderInfo1.itemnumber = "P1001";
        purchaseOrderInfo1.itemdescription = "Pencils HB";
        purchaseOrderInfo1.quantity = 5;
        purchaseOrderInfo1.um = "dozen";
        purchaseOrderInfo1.price = 10;
        purchaseOrderInfoList.add(purchaseOrderInfo1);

        PurchaseOrderInfo purchaseOrderInfo2 = new PurchaseOrderInfo();
        purchaseOrderInfo2.s_no = 2;
        purchaseOrderInfo2.itemnumber = "P1003";
        purchaseOrderInfo2.itemdescription = "Pencils 2B";
        purchaseOrderInfo2.quantity = 4;
        purchaseOrderInfo2.um = "dozen";
        purchaseOrderInfo2.price = 10;
        purchaseOrderInfoList.add(purchaseOrderInfo2);

        PurchaseOrderInfo purchaseOrderInfo3 = new PurchaseOrderInfo();
        purchaseOrderInfo3.s_no = 3;
        purchaseOrderInfo3.itemnumber = "P1003";
        purchaseOrderInfo3.itemdescription = "Paper A4 - Photo Copier";
        purchaseOrderInfo3.quantity = 10;
        purchaseOrderInfo3.um = "ream";
        purchaseOrderInfo3.price = 3;
        purchaseOrderInfoList.add(purchaseOrderInfo3);

        PurchaseOrderInfo purchaseOrderInfo4 = new PurchaseOrderInfo();
        purchaseOrderInfo4.s_no = 4;
        purchaseOrderInfo4.itemnumber = "P1234";
        purchaseOrderInfo4.itemdescription = "Pens - Ball point";
        purchaseOrderInfo4.quantity = 15;
        purchaseOrderInfo4.um = "boxes";
        purchaseOrderInfo4.price = 2;
        purchaseOrderInfoList.add(purchaseOrderInfo4);

        PurchaseOrderInfo purchaseOrderInfo5 = new PurchaseOrderInfo();
        purchaseOrderInfo5.s_no = 5;
        purchaseOrderInfo5.itemnumber = "P3221";
        purchaseOrderInfo5.itemdescription = "Highligter";
        purchaseOrderInfo5.quantity = 8;
        purchaseOrderInfo5.um = "sets";
        purchaseOrderInfo5.price = 10;
        purchaseOrderInfoList.add(purchaseOrderInfo5);

        PurchaseOrderBasicInfo orderbasicInfo = new PurchaseOrderBasicInfo();
        orderbasicInfo.iD = "US120499";
        orderbasicInfo.orderDate = new GregorianCalendar(2019, 7, 7).getTime();
        orderbasicInfo.creditTerms = "30";
        orderbasicInfo.pONumber = "PO1011";
        orderbasicInfo.ref = "QT1231";
        orderbasicInfo.deliverToCompany = "Sanfort Pvt. Ltd.";
        orderbasicInfo.deliverToAddress = "1322, High Street, Geln Waverlay";
        orderbasicInfo.postalCode = "Victoria 3456";
        orderbasicInfo.country = "Australia";
        ///#endregion

        //Add data source
        workbook.addDataSource("po", purchaseOrderInfoList);
        workbook.addDataSource("tax", 5);
        workbook.addDataSource("ds", orderbasicInfo);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_PurchaseOrder.xlsx";
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
        return new String[] { "xlsx/Template_PurchaseOrder.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "PurchaseOrderInfo", "PurchaseOrderBasicInfo" };
	}
}