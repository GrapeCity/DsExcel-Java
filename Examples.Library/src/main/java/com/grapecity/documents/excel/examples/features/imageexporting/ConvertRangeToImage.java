package com.grapecity.documents.excel.examples.features.imageexporting;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertRangeToImage extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1:C1").setValue(new String[] { "Device", "Quantity", "Unit Price" });
        worksheet.getRange("A2:C5").setValue(new Object[][]
        {
            { "T540p", 12, 9850 },
            { "T570", 5, 7460 },
            { "Y460", 6, 5400 },
            { "Y460F", 8, 6240 }
         });

        // Set style
        worksheet.getRange("A1:C1").getFont().setBold(true);
        worksheet.getRange("A1:C1").getFont().setColor(Color.GetWhite());
        worksheet.getRange("A1:C1").getInterior().setColor(Color.GetLightBlue());
        worksheet.getRange("A2:C5").getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetOrange());
        worksheet.getRange("A2:C5").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.DashDot);

        // Save the range "A1:C5" as an image to a stream.
        worksheet.getRange("A1:C5").toImage(outputStream, ImageType.PNG);
	}

	@Override
	public boolean getSaveAsImage() {
        return true;
	}
}
