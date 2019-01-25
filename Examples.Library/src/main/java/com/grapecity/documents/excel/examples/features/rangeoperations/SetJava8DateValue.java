package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.util.GregorianCalendar;

public class SetJava8DateValue extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange a1 = worksheet.getRange("A1");

        //Java 8 introduces a new package java.time which contains lots of new date/time types and sub-packages to support JSR310.
        //GcExcel can handle these new types when working with Java 8 or upper.
        a1.setValue(java.time.LocalDateTime.now());
        java.time.LocalDateTime java8Date = (java.time.LocalDateTime)a1.getValue();

        a1.setNumberFormat("m/d/yyyy h:mm");
        a1.setColumnWidth(20);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
