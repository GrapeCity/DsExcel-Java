package com.grapecity.documents.excel.examples.features.rangeoperations;

import java.util.EnumSet;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PasteType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CopyPasteOptions extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set data of PC
        worksheet.getRange("A2").setValue("PC");
        worksheet.getRange("A4:C4").setValue(new String[] {"Device", "Quantity", "Unit Price"});
        worksheet.getRange("A5:C10").setValue(new Object[][]
        {
        	{"T540p", 12, 9850},
        	{"T570", 5, 7460},
        	{"Y460", 6, 5400},
        	{"Y460F", 8, 6240}
        });

        //Set style
        worksheet.getRange("A2").setRowHeight(30);
        worksheet.getRange("A2").getFont().setSize(20);
        worksheet.getRange("A2").getFont().setBold(true);
        worksheet.getRange("A4:C4").getFont().setBold(true);
        worksheet.getRange("A4:C4").getFont().setColor(Color.GetWhite());
        worksheet.getRange("A4:C4").getInterior().setColor(Color.GetLightBlue());
        worksheet.getRange("A5:C10").getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetOrange());
        worksheet.getRange("A5:C10").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.DashDot);

        //Copy only style and row height
        worksheet.getRange("H1").setValue("Copy style and row height from previous cells.");
        worksheet.getRange("H1").getFont().setColor(Color.GetRed());
        worksheet.getRange("H1").getFont().setBold(true);
        worksheet.getRange("A2:C10").copy(worksheet.getRange("H2"), EnumSet.of(PasteType.Formats));

        //Set data of mobile devices
        worksheet.getRange("H2").setValue("Mobile");
        worksheet.getRange("H4:J4").setValue(new String[] {"Device", "Quantity", "Unit Price"});
        worksheet.getRange("H5:J10").setValue(new Object[][]
        {
        	{"HW-P30", 20, 4200},
        	{"IPhone-X", 5, 9888},
        	{"IPhone-6s plus", 15, 6880}
        });

        //Add new sheet
        IWorksheet worksheet2 =workbook.getWorksheets().add();

        //Copy only style to new sheet
        worksheet.getRange("A2:C10").copy(worksheet2.getRange("A2"), EnumSet.of(PasteType.Formats));
        worksheet2.getRange("A3").setValue("Copy style from sheet1.");
        worksheet2.getRange("A3").getFont().setColor(Color.GetRed());
        worksheet2.getRange("A3").getFont().setBold(true);
    }
    
}
