package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AllBarcodes extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("A:A").setColumnWidth(2);
		worksheet.getRange("B:C").setColumnWidth(15);
		worksheet.getRange("D:G").setColumnWidth(25);
		worksheet.getRange("4:14").setRowHeight(57);
		worksheet.getRange("B3").setValue("Type");
		worksheet.getRange("C3").setValue("Data");
		worksheet.getRange("B2").setValue("Barcode");
		worksheet.getRange("B2:G2").setMergeCells(true);
		worksheet.getRange("D3:G3").setValue(new Object[][]
		{
			{"Default", "Change color", "Change showLabel", "Change labelPosition"}
		});
		worksheet.getPageSetup().setPrintTitleColumns("$A:$C");
		worksheet.getRange("B4:C14").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C14").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:G3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:G3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C14").setValue(new Object[][]
		{
			{"QR code", "Policy:411"},
			{"Data Matrix", "Policy:411"},
			{"PDF417", "6935205311092"},
			{"EAN-8", "4137962"},
			{"EAN-13", "6920312296219"},
			{"Code39", "3934712708295"},
			{"Code93", "6945091701532"},
			{"Code49", "6901668002433"},
			{"Code128", "465465145645"},
			{"Codabar", "9787560044231"},
			{"gs1128", "010123456789012815051231"}
		});
		String[] types = {"BC_QRCODE", "BC_DataMatrix", "BC_PDF417", "BC_EAN8", "BC_EAN13", "BC_CODE39", "BC_CODE93", "BC_CODE49", "BC_CODE128", "BC_CODABAR", "BC_GS1_128"};
		worksheet.getPageSetup().setPrintGridlines(true);

		//use formula to add barcode
		for (int i = 0; i < types.length; i++)
		{
			String columnD = "D" + (i + 4);
			String columnE = "E" + (i + 4);
			worksheet.getRange(columnD).setFormula("=" + types[i] + "(C" + (i + 4) + ")");
			worksheet.getRange(columnE).setFormula("=" + types[i] + "(C" + (i + 4) + ",\"#fff\",\"#000\")");
		}

		for (int i = 3; i < types.length; i++)
		{
			String columnD = "F" + (i + 4);
			String columnE = "G" + (i + 4);
			worksheet.getRange(columnD).setFormula("=" + types[i] + "(C" + (i + 4) + ",,,0)");
			worksheet.getRange(columnE).setFormula("=" + types[i] + "(C" + (i + 4) + ",,,,\"top\")");
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