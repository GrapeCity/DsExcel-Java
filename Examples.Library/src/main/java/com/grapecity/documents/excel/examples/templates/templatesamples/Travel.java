package com.grapecity.documents.excel.examples.templates.templatesamples;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.GregorianCalendar;
import java.util.List;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Travel extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//Load template file Template_Score.xlsx from resource
		InputStream templateFile = this.getResourceStream("xlsx/Template_Travel.xlsx");
		workbook.open(templateFile);

		/// #region Define custom classes
		//		public class FlightInfo {
		//			public String carrier;
		//			public int flightNo;
		//			public Date date = new Date(0);
		//			public String from;
		//			public Calendar departureTime;
		//			public String to;
		//			public Calendar arrivalTime;
		//			public String reservationNo;
		//		}

		//		public class LoadingInfo {
		//			public String accommodations;
		//			public Date date = new Date(0);
		//			public String concierge;
		//			public String phone;
		//			public String email;
		//			public String addressPart1;
		//			public String addressPart2;
		//			public String confirmNo;
		//			public int days;
		//			public double totalCost;
		//		}

		//		public class ContactInfo {
		//			public String contact;
		//			public String phone;
		//			public String notes;
		//		}
		/// #endregion

		///#region Init Data
		List<FlightInfo> flightInfos = new ArrayList<FlightInfo>();

		FlightInfo flightInfo1 = new FlightInfo();
		flightInfo1.carrier = "Trenz Airlines";
		flightInfo1.flightNo = 1623;
		flightInfo1.date = new GregorianCalendar(2018, 10, 25).getTime();
		flightInfo1.from = "Lorem International";

		GregorianCalendar departureTime1 = new GregorianCalendar();
		departureTime1.clear();
		departureTime1.set(0, 0, 0, 7, 56, 0);
		flightInfo1.departureTime = departureTime1;

		flightInfo1.to = "Dolor Airport";

		GregorianCalendar arrivalTime1 = new GregorianCalendar();
		arrivalTime1.clear();
		arrivalTime1.set(0, 0, 0, 9, 15, 0);
		flightInfo1.arrivalTime = arrivalTime1;

		flightInfo1.reservationNo = "AG4567997";
		flightInfos.add(flightInfo1);

		FlightInfo flightInfo2 = new FlightInfo();
		flightInfo2.carrier = "Trenz Airlines";
		flightInfo2.flightNo = 1323;
		flightInfo2.date = new GregorianCalendar(2018, 10, 30).getTime();
		flightInfo2.from = "Lorem International";

		GregorianCalendar departureTime2 = new GregorianCalendar();
		departureTime2.clear();
		departureTime2.set(0, 0, 0, 20, 25, 0);
		flightInfo2.departureTime = departureTime2;

		flightInfo2.to = "Dolor Airport";

		GregorianCalendar arrivalTime2 = new GregorianCalendar();
		arrivalTime2.clear();
		arrivalTime2.set(0, 0, 0, 21, 45, 0);
		flightInfo2.arrivalTime = arrivalTime2;

		flightInfo2.reservationNo = "AG4567998";
		flightInfos.add(flightInfo2);

		List<LoadingInfo> loadingInfos = new ArrayList<LoadingInfo>();

		LoadingInfo loadingInfo1 = new LoadingInfo();
		loadingInfo1.accommodations = "Lorem Hotel";
		loadingInfo1.date = new GregorianCalendar(2018, 10, 25).getTime();
		loadingInfo1.concierge = "Charles";
		loadingInfo1.phone = "01234 567 890";
		loadingInfo1.email = "charles@lorem.com";
		loadingInfo1.addressPart1 = "123 High Street, ";
		loadingInfo1.addressPart2 = "Anytown, County, Postcode";
		loadingInfo1.confirmNo = "A4567";
		loadingInfo1.days = 2;
		loadingInfo1.totalCost = 800;
		loadingInfos.add(loadingInfo1);

		LoadingInfo loadingInfo2 = new LoadingInfo();
		loadingInfo2.accommodations = "Deloz Hotel";
		loadingInfo2.date = new GregorianCalendar(2018, 10, 27).getTime();
		loadingInfo2.concierge = "James";
		loadingInfo2.phone = "01234 567 890";
		loadingInfo2.email = "no_reply@example.com";
		loadingInfo2.addressPart1 = "202 Halford Street, ";
		loadingInfo2.addressPart2 = "Anytown, County, Postcode";
		loadingInfo2.confirmNo = "A4568";
		loadingInfo2.days = 3;
		loadingInfo2.totalCost = 900;
		loadingInfos.add(loadingInfo2);

		List<ContactInfo> emergencyContactInfos = new ArrayList<ContactInfo>();

		ContactInfo emergencyContactInfo1 = new ContactInfo();
		emergencyContactInfo1.contact = "Airline Reservations";
		emergencyContactInfo1.phone = "01234 567 890";
		emergencyContactInfos.add(emergencyContactInfo1);

		ContactInfo emergencyContactInfo2 = new ContactInfo();
		emergencyContactInfo2.contact = "Hotel Reservations";
		emergencyContactInfo2.phone = "12342322232";
		emergencyContactInfos.add(emergencyContactInfo2);

		List<ContactInfo> contactInfos = new ArrayList<ContactInfo>();

		ContactInfo contactInfo1 = new ContactInfo();
		contactInfo1.contact = "Tom Jenkins";
		contactInfo1.phone = "01234 567 890";
		contactInfo1.notes = "tom.jerkins@trenz.com";
		contactInfos.add(contactInfo1);

		ContactInfo contactInfo2 = new ContactInfo();
		contactInfo2.contact = "Rayna James";
		contactInfo2.phone = "19234222456";
		contactInfo1.notes = "ratna.james@deloz.com";
		contactInfos.add(contactInfo2);
		///#endregion

		//Add data source
		workbook.addDataSource("ds1", flightInfos);
		workbook.addDataSource("ds2", loadingInfos);
		workbook.addDataSource("ds3", emergencyContactInfos);
		workbook.addDataSource("ds4", contactInfos);
		//Invoke to process the template
		workbook.processTemplate();
	}

	@Override
	public boolean getIsNew() {
		return true;
	}

	@Override
	public String getTemplateName() {
		return "Template_Travel.xlsx";
	}

	@Override
	public boolean getHasTemplate() {
		return true;
	}

	@Override
	public String[] getResources() {
		return new String[] { "xlsx/Template_Travel.xlsx" };
	}
	
	@Override
	public String[] getRefs() {
		return new String[] { "FlightInfo", "LoadingInfo", "ContactInfo" };
	}
}