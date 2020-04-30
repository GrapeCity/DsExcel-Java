package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.GradientStyle;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.PathShapeType;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.InputStream;

public class ConfigGradientFillWithPathShape extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/WebsiteFlowChart.xlsx");
        workbook.open(fileStream);

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Get "Idea"
        IShape idea = worksheet.getShapes().get("Idea");
        idea.getFill().twoColorGradient(GradientStyle.FromCenter, 1);

        idea.getFill().getGradientStops().get(0).setPosition(0.33);
        idea.getFill().getGradientStops().get(0).getColor().setRGB(Color.FromArgb(0, 112, 192));

        idea.getFill().getGradientStops().get(1).setPosition(1);
        idea.getFill().getGradientStops().get(1).getColor().setRGB(Color.GetWhite());

        //Set gradient path type as "Path"
        idea.getFill().setGradientPathType(PathShapeType.Path);

        //Get "Functionality"
        IShape functionality = worksheet.getShapes().get("Functionality");
        functionality.getFill().twoColorGradient(GradientStyle.FromCenter, 1);

        functionality.getFill().getGradientStops().get(0).setPosition(0.33);
        functionality.getFill().getGradientStops().get(0).getColor().setRGB(Color.FromArgb(0, 112, 192));

        functionality.getFill().getGradientStops().get(1).setPosition(1);
        functionality.getFill().getGradientStops().get(1).getColor().setRGB(Color.GetWhite());

        //Set gradient path type as "Path"
        functionality.getFill().setGradientPathType(PathShapeType.Path);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }

    @Override
    public boolean getShowScreenshot() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/WebsiteFlowChart.xlsx" };
    }
}
