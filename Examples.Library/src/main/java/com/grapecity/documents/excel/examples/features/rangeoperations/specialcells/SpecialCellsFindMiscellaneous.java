package com.grapecity.documents.excel.examples.features.rangeoperations.specialcells;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SpecialCellsFindMiscellaneous extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet ws = workbook.getActiveSheet();

        // Prepare data
        Object[][] rngA1D2 = new Object[][] { { "Register", null, null, null },
                { "Field name", "Wildcard", "Validation error", "User input" } };
        ws.getRange("$A$1:$D$2").setValue(rngA1D2);

        Object[][] rngA3C6 = new Object[][] { { "User name", "??*", "At least 2 characters" },
                { "Captcha", "?????", "5 characters required" }, { "E-mail", "?*@?*.?*", "The format is incorrect" },
                { "Security code", "#######", "7 digits required" } };
        ws.getRange("$A$3:$C$6").setValue(rngA3C6);

        Object[][] rngA8D14 = new Object[][] { { "User table", null, null, null }, { "Id", "Name", "Email", "Banned" },
                { 1d, "User 1", "8zgnvlkp2@163.com", true }, { 2d, "User 2", "b9fvaswb@163.com", false },
                { 3d, "User", "md78b", false }, { 4d, "User 4", "1qasghjfg@163.com", false },
                { 5d, "U", "mncx23k8@163.com", false } };
        ws.getRange("$A$8:$D$14").setValue(rngA8D14);

        ws.getRange("A1:D1").merge();
        ws.getRange("A1:D1").setHorizontalAlignment(HorizontalAlignment.Center);
        ws.getRange("A8:D8").merge();
        ws.getRange("A8:D8").setHorizontalAlignment(HorizontalAlignment.Center);

        ws.getRange("D3").addComment("Required");
        ws.getRange("D4").addComment("Required");
        ws.getRange("D5").addComment("Required");
        ws.getRange("D6").addComment("Required");

        ws.getRange("D10:D14").getValidation().add(ValidationType.List, ValidationAlertStyle.Stop,
                ValidationOperator.Between, "True,False", null);

        IFormatCondition condition = (IFormatCondition) ws.getRange("C10:C14").getFormatConditions().add(
                FormatConditionType.Expression, FormatConditionOperator.Between, "=ISERROR(MATCH($B$5,C10,0))", null);
        condition.getFont().setColor(Color.GetRed());

        IFormatCondition condition2 = (IFormatCondition) ws.getRange("B10:B14").getFormatConditions()
                .add(FormatConditionType.Expression, FormatConditionOperator.Between, "=LEN(B10)<=2", null);
        condition2.getFont().setColor(Color.GetRed());

        ws.getRange("4:4").getEntireRow().setHidden(true);

        IRange searchScope = ws.getRange("1:14");

        // Find comments
        IRange comments = searchScope.specialCells(SpecialCellType.Comments);

        // Find last cell
        IRange lastCell = searchScope.specialCells(SpecialCellType.LastCell);

        // Find visible
        IRange visible = searchScope.specialCells(SpecialCellType.Visible);

        // Find blanks
        IRange blanks = searchScope.specialCells(SpecialCellType.Blanks);

        // Find all format conditions
        IRange allFormatConditions = searchScope.specialCells(SpecialCellType.AllFormatConditions);

        // Find all validation
        IRange allValidation = searchScope.specialCells(SpecialCellType.AllValidation);

        // Find same format condition as B10
        IRange sameFormatConditions = ws.getRange("B10").specialCells(SpecialCellType.SameFormatConditions);

        // Find same validation as D10
        IRange sameValidation = ws.getRange("D10").specialCells(SpecialCellType.SameValidation);

        // Find merged cells
        IRange merged = searchScope.specialCells(SpecialCellType.MergedCells);

        // Output
        ws.getRange("A16").setValue("Find result");
        ws.getRange("A16:C16").merge();
        ws.getRange("A16:C16").setHorizontalAlignment(HorizontalAlignment.Center);
        ws.getRange("$A$17:$A$25")
                .setValue(new Object[][] { { "Comments" }, { "LastCell" }, { "Visible" }, { "Blanks" },
                        { "AllFormatConditions" }, { "AllValidation" }, { "SameFormatConditions B10" },
                        { "SameValidation D10" }, { "MergedCells" } });
        ws.getRange("$C$17:$C$25")
                .setValue(new Object[][] { { comments.getAddress() }, { lastCell.getAddress() },
                        { visible.getAddress() }, { blanks.getAddress() }, { allFormatConditions.getAddress() },
                        { allValidation.getAddress() }, { sameFormatConditions.getAddress() },
                        { sameValidation.getAddress() }, { merged.getAddress() } });

        ws.getUsedRange().getEntireColumn().autoFit();
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
