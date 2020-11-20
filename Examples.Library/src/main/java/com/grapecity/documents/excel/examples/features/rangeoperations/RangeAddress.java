package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class RangeAddress extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IRange mc = workbook.getWorksheets().get("Sheet1").getCells().get(0, 0);
        System.out.println(mc.getAddress()); // $A$1
        System.out.println(mc.getAddress(false, true)); // $A1
        System.out.println(mc.getAddress(true, true, ReferenceStyle.R1C1)); // R1C1
        System.out.println(
                mc.getAddress(false, false, ReferenceStyle.R1C1, workbook.getWorksheets().get(0).getCells().get(2, 2))); // R[-2]C[-2]
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
    public boolean getCanDownload() {
        return false;
    }
}
