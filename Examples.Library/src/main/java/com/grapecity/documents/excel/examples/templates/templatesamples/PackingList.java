package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.util.ArrayList;
import java.util.Date;

public class PackingList {
	public String exporter;
	public String address_exporter;
	public String country_exporter;
	public String phonenumber_shipper;
	public String shipper;

	public String imports;
	public String address_consignee;
	public String country_consignee;
	public String phonenumber_consignee;
	public String consignee;

	public int invoice_No;
	public Date date = new Date(0);
	public int reference;

	public String dispatchMethod;
	public String shipmentType;
	public String vA;
	public String voyageNo;
	public String portofLanding;
	public Date departureDate = new Date(0);
	public String dischargePort;
	public String finalDestination;

	public String goodsOriginCountry;
	public String destinationCountry;

	public ArrayList<Product> item;

	public String issuePlace;
	public Date issueDate = new Date(0);
	public String signatoryCompany;
	public String signatoryName;

}