package com.grapecity.documents.excel.examples.features.barcode;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Code49 extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
		//set worksheet layout and data
		IWorksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getRange("B:C").setColumnWidth(10);
		worksheet.getRange("D:F").setColumnWidth(30);
		worksheet.getRange("4:5").setRowHeight(80);
		worksheet.getRange("A:A").setColumnWidth(5);
		worksheet.getRange("B2").setValue("Code49");
		worksheet.getRange("B2:F2").setMergeCells(true);
		worksheet.getRange("B3:G3").setValue(new Object[][]
		{
			{"Name", "Number", "Default", "Customer Label Font", "Line Through Label"}
		});
		worksheet.getRange("B4:C7").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B4:C7").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B2:F3").setHorizontalAlignment(HorizontalAlignment.Center);
		worksheet.getRange("B2:F3").setVerticalAlignment(VerticalAlignment.Center);
		worksheet.getRange("B4:C5").setValue(new Object[][]
		{
			{"Police", "911"},
			{"Travel Info Call 511", "511"}
		});
		worksheet.getRange("B4:C6").setWrapText(true);
		worksheet.getRange("G6").setWrapText(true);
		worksheet.getPageSetup().setPrintGridlines(true);
		worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);

		//set formula
		for (int i = 4; i < 6; i++)
		{
			String value = "CONCAT(B" + i + ", \": \",C" + i + ")";
			worksheet.getRange("D" + i).setFormula("=BC_CODE49" + "(" + value + ")");
			worksheet.getRange("E" + i).setFormula("=BC_CODE49" + "(" + value + ", , , true, \"top\", false, 0, \"Arial\", \"normal\", 700)");
			worksheet.getRange("F" + i).setFormula("=BC_CODE49" + "(" + value + ", , , , , , , , , 700, \"line - through\", \"left\", \"24px\")");
		}
	}

	@Override
	public boolean getCanDownload() {
		return false;
	}

	@Override
	public boolean getIsNew()
	{
		return true;
	}
}