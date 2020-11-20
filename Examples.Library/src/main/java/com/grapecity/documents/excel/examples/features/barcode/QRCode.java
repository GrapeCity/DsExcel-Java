package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class QRCode extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:K").setColumnWidth(15);
		worksheet.getRange("4:6").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(2);
		worksheet.getRange("B2").setValue("QR Code");
		worksheet.getRange("B2:K2").setMergeCells(true);
		worksheet.getRange("I3:J3").setMergeCells(true);
		worksheet.getRange("B3:H3").setValue(new Object[][]
		{
			{"Server", "Data", "Default", "Change errorCorrectionLevel", "Change model", "Change version", "Change mask"}
		});
		worksheet.getRange("I3").setValue("Change connection and connectionNo");
		worksheet.getRange("K3:K5").setValue(new Object[][]
		{
			{"Explain"},
			{"No QR Code generated, barcode data is too short to create connection symbol."},
			{"No QR Code generated, barcode data is too short to create connection symbol."}
		});
		worksheet.getPageSetup().setPrintTitleColumns("$A:$C");
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);
		worksheet.getRange("K4:K5").getFont().setColor(Color.GetRed());
		worksheet.getRange("K4:K5").setWrapText(true);
		worksheet.getRange("B4:C6").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C6").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:K3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:K3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C6").setValue(new Object[][]
		{
			{"Police", "911"},
			{"Travel Info Call 511", "511"},
			{"", "www.grapecity.com"}
		});
		//set formula
		for (int i = 4; i < 7; i++)
		{
			worksheet.getRange("D" + i).setFormula("=BC_QRCODE" + "(C" + i + ")");
			worksheet.getRange("E" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,\"H\")");
			worksheet.getRange("F" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,,1)");
			worksheet.getRange("G" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,,,8)");
			worksheet.getRange("H" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,,,,3)");
			worksheet.getRange("I" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,,,,,\"true\",0)");
			worksheet.getRange("J" + i).setFormula("=BC_QRCODE" + "(C" + i + ",,,,,,,\"true\",1)");
		}
	}
	@Override
	public boolean getSavePdf()
	{
		return true;
	}
	@Override
	public boolean getIsNew()
	{
		return true;
	}
}