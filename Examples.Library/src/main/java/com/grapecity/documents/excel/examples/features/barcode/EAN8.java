package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class EAN8 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:G").setColumnWidth(17);
		worksheet.getRange("4:7").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("EAN-8");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:G3").setValue(new Object[][]
		{
			{"Name", "Number", "Default", "Change showLabel", "Change labelPosition", "Explain"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C6").setValue(new Object[][]
		{
			{"Value length is 7", "4137962"},
			{"Value length is 8", "81424863"},
			{"value length is 8", "81424865"}
		});
		worksheet.getRange("G6").setValue("No EAN-8 generated, because the last digit is check-sum digit and it is invalid");
		worksheet.getRange("G6").getFont().setColor(Color.GetRed());
		worksheet.getRange("B4:C6").setWrapText(true);
		worksheet.getRange("G6").setWrapText(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);
		//set formula
		for (int i = 4; i < 7; i++)
		{
			worksheet.getRange("D" + i).setFormula("=BC_EAN8" + "(C" + i + ")");
			worksheet.getRange("E" + i).setFormula("=BC_EAN8" + "(C" + i + ",,,0)");
			worksheet.getRange("F" + i).setFormula("=BC_EAN8" + "(C" + i + ",,,,\"top\")");
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