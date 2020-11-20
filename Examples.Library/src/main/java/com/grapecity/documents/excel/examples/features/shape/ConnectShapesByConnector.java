package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.ConnectorType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConnectShapesByConnector extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shapeBegin = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 1, 1, 100, 100);
        IShape endBegin = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 200, 200, 100, 100);
        IShape ConnectorShape = worksheet.getShapes().addConnector(ConnectorType.Straight, 1, 1, 101, 101);

        //connect shapes by connector shape.
        ConnectorShape.getConnectorFormat().beginConnect(shapeBegin, 3);
        ConnectorShape.getConnectorFormat().endConnect(endBegin, 0);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

}
