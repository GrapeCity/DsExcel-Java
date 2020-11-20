package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Codabar extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:F").setColumnWidth(20);
		worksheet.getRange("4:7").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("Codabar");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:G3").setValue(new Object[][]
		{
			{"Name", "Number", "Default", "Change checkDigit", "Change nwRatio"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C6").setValue(new Object[][]
		{
			{"Notebook", "6935205311092"},
			{"Paper", "6922266446146"},
			{"Value can contain letters and some symbol", "A1234+-/.$A"}
		});
		worksheet.getRange("B4:C6").setWrapText(true);
		worksheet.getRange("G6").setWrapText(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);

		//set formula
		for (int i = 4; i < 7; i++)
		{
			worksheet.getRange("D" + i).setFormula("=BC_CODABAR" + "(C" + i + ")");
			worksheet.getRange("E" + i).setFormula("=BC_CODABAR" + "(C" + i + ",,,,,\"true\")");
			worksheet.getRange("F" + i).setFormula("=BC_CODABAR" + "(C" + i + ",,,,,,\"2\")");
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