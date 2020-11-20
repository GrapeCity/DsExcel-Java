package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class EAN13 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:C").setColumnWidth(15);
		worksheet.getRange("D:G").setColumnWidth(20);
		worksheet.getRange("4:7").setRowHeight(60);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("EAN-13");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:G3").setValue(new Object[][]
		{
			{"Name", "Number", "Default", "Change addOn", "Change addOnLabelPosition", "Explain"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C6").setValue(new Object[][]
		{
			{"Medicine", "692031229621"},
			{"Pen", "6945091701532"},
			{"value length is 13", "8142486545683"}
		});
		worksheet.getRange("G6").setValue("No EAN-13 generated, because the last digit is check-sum digit and it is invalid");
		worksheet.getRange("G6").getFont().setColor(Color.GetRed());
		worksheet.getRange("B4:C6").setWrapText(true);
		worksheet.getRange("G6").setWrapText(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		worksheet.getPageSetup().setPrintGridlines(true);
		//set formula
		for (int i = 4; i < 7; i++)
		{
			worksheet.getRange("D" + i).setFormula("=BC_EAN13" + "(C" + i + ")");
			worksheet.getRange("E" + i).setFormula("=BC_EAN13" + "(C" + i + ",,,,,22)");
			worksheet.getRange("F" + i).setFormula("=BC_EAN13" + "(C" + i + ",,,,,22,\"bottom\")");
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