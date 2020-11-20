package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Code128 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:F").setColumnWidth(20);
		worksheet.getRange("4:7").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("Code128");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:F3").setValue(new Object[][]
		{
			{"Name", "Number", "Default", "Hidden Label", "Custom Label Font"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C7").setValue(new Object[][]
		{
			{"Police", 911},
			{"Telephone Directory Assistance", 411},
			{"Non-emergency Municipal Services", 311},
			{"Travel Info Call 511", 511}
		});
		worksheet.getRange("B4:C6").setWrapText(true);
		worksheet.getRange("G6").setWrapText(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);
		//set formula
		for (int i = 4; i < 8; i++)
		{
			worksheet.getRange("D" + i).setFormula("=BC_CODE128" + "(C" + i + ")");
			worksheet.getRange("E" + i).setFormula("=BC_CODE128" + "(C" + i + ", , , false))");
			worksheet.getRange("F" + i).setFormula("=BC_CODE128" + "(C" + i + ", , , true, \"top\", \"B\",\"Arial\", \"normal\")");
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