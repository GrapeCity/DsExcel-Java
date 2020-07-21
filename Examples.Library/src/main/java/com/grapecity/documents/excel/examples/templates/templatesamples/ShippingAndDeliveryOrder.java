package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.GregorianCalendar;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShippingAndDeliveryOrder extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        //Load template file from resource
        InputStream templateFile = this.getResourceStream("xlsx/Template_ShippingAndDeliveryOrder.xlsx");
        workbook.open(templateFile);

        ///#region Define custom classes
        //        public class PackingList {
        //        	public String exporter;
        //        	public String address_exporter;
        //        	public String country_exporter;
        //        	public String phonenumber_shipper;
        //        	public String shipper;
        //
        //        	public String imports;
        //        	public String address_consignee;
        //        	public String country_consignee;
        //        	public String phonenumber_consignee;
        //        	public String consignee;
        //
        //        	public int invoice_No;
        //        	public Date date = new Date(0);
        //        	public int reference;
        //
        //        	public String dispatchMethod;
        //        	public String shipmentType;
        //        	public String vA;
        //        	public String voyageNo;
        //        	public String portofLanding;
        //        	public Date departureDate = new Date(0);
        //        	public String dischargePort;
        //        	public String finalDestination;
        //
        //        	public String goodsOriginCountry;
        //        	public String destinationCountry;
        //
        //        	public ArrayList<Product> item;
        //
        //        	public String issuePlace;
        //        	public Date issueDate = new Date(0);
        //        	public String signatoryCompany;
        //        	public String signatoryName;
        //
        //        }

        //        public class Product {
        //        	public String productcode;
        //        	public String goods;
        //        	public double quantity;
        //        	public double netweight;
        //        	public String kindAndPackagesCount;
        //        	public double grossweight;
        //        	public double measurements;
        //        }
        ///#endregion

        ///#region Init Data
        PackingList packinginfo = new PackingList();
        packinginfo.exporter = "DEL Exports";
        packinginfo.address_exporter = "4243 Longline Vlvd Longline, CA - 98020";
        packinginfo.country_exporter = "United States";
        packinginfo.phonenumber_shipper = "010-510-22424";
        packinginfo.shipper = "Diana Thompson";

        packinginfo.imports = "Deanna Imports";
        packinginfo.address_consignee = "113/23, Lombard Street Halford Townsville, Melbourne, 4323";
        packinginfo.country_consignee = "Australia";
        packinginfo.phonenumber_consignee = "010-510-33232";
        packinginfo.consignee = "James Williams";

        packinginfo.invoice_No = 1934;
        packinginfo.date = new GregorianCalendar(2019, 1, 30).getTime();
        packinginfo.reference = 1934;

        packinginfo.dispatchMethod = "Sea";
        packinginfo.shipmentType = "FCL";
        packinginfo.goodsOriginCountry = "United States";
        packinginfo.destinationCountry = "Australia";
        packinginfo.vA = "MAKERS DYER";
        packinginfo.voyageNo = "6E";
        packinginfo.portofLanding = "Longline - California";
        packinginfo.departureDate = new GregorianCalendar(2019, 2, 1).getTime();
        packinginfo.dischargePort = "Melbourne - Australia";
        packinginfo.finalDestination = "Australia";

        Product product1 = new Product();
        product1.productcode = "P1001";
        product1.goods = "Pencils - HB";
        product1.quantity = 5;
        product1.netweight = 0.1;
        product1.kindAndPackagesCount = "PALLET X 1";
        product1.grossweight = 750;
        product1.measurements = 1.7;

        Product product2 = new Product();
        product2.productcode = "P1002";
        product2.goods = "Paper - A4";
        product2.quantity = 3;
        product2.netweight = 2;
        product2.kindAndPackagesCount = "PALLET X 2";
        product2.grossweight = 250;
        product2.measurements = 1.4;

        packinginfo.item = new ArrayList<Product>(Arrays.asList(product1, product2));

        packinginfo.issuePlace = "Longline";
        packinginfo.issueDate = new GregorianCalendar(2019, 1, 30).getTime();
        packinginfo.signatoryCompany = "DEL Exports";
        packinginfo.signatoryName = "Rayna Johnson";
        ///#endregion

        //Add data source
        workbook.addDataSource("ds", packinginfo);
        //Invoke to process the template
        workbook.processTemplate();
	}

	@Override
	public String getTemplateName() {
        return "Template_ShippingAndDeliveryOrder.xlsx";
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
        return new String[] { "xlsx/Template_ShippingAndDeliveryOrder.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
        return new String[] { "PackingList", "Product" };
	}
}