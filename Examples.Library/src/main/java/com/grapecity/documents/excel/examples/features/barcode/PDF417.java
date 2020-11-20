package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PDF417 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:C").setColumnWidth(12);
		worksheet.getRange("D:F").setColumnWidth(30);
		worksheet.getRange("4:7").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("Data Matrix");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:F3").setValue(new Object[][]
		{
			{"Server", "Data", "Default", "Customer Padding", "Customer Columns Count"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C7").setValue(new Object[][]
		{
			{"Police", "911"},
			{"Telephone Directory Assistance", "411"},
			{"Non-emergency Municipal Services", "311"},
			{"Travel Info Call 511", "511"}
		});
		worksheet.getRange("B4:B7").setWrapText(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);
		//set formula
		for (int i = 4; i < 8; i++)
		{
			String value = "CONCAT(B" + i + ",\":\",C" + i + ")";
			worksheet.getRange("D" + i).setFormula("=BC_PDF417" + "(" + value + ")");
			worksheet.getRange("E" + i).setFormula("=BC_PDF417" + "(" + value + ", , , , , , , 0, 10, 5, 5)");
			worksheet.getRange("F" + i).setFormula("=BC_PDF417" + "(" + value + ", , , , , 5)");
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