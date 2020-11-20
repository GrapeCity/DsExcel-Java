package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.HorizontalAnchor;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.VerticalAnchor;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeWithTextAlignment extends ExampleBase {
	@Override
	public void execute(Workbook workbook){
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/WebsiteFlowChart[Template].xlsx");
        workbook.open(fileStream);
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        for (IShape shape : worksheet.getShapes())
        {
            //Centers text vertically. 
            shape.getTextFrame().setVerticalAnchor(VerticalAnchor.AnchorMiddle);
            //Centers text horizontally.
            shape.getTextFrame().setHorizontalAnchor(HorizontalAnchor.Center);
        }
	}
	
	@Override
    public boolean getShowViewer()
    {
        return false;
    }

	@Override
    public boolean getSavePdf()
    {
         return true;
    }

	@Override
	public String[] getResources() {
        return new String[] { "xlsx/WebsiteFlowChart[Template].xlsx"};
	}
}
