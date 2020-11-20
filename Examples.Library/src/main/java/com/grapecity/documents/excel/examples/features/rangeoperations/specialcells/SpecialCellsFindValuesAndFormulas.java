package com.grapecity.documents.excel.examples.features.rangeoperations.specialcells;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SpecialCellsFindValuesAndFormulas extends ExampleBase {

	@Override
	public void execute(Workbook workbook) {
		IWorksheet ws = workbook.getActiveSheet();

		// Prepare data
		ws.getRange("$A$1:$F$1").setValue(new Object[][] {
				{ "Test id", "Group id", "Group item id", "New test id", "Test result", "Error code" } });

		ws.getRange("$B$2:$C$2").setValue(1d);
		ws.getRange(
				"$E$2,$E$7,$E$12,$E$21,$E$27,$E$36,$E$40,$E$47:$E$48,$E$51,$E$59:$E$60,$E$70:$E$71,$E$80:$E$81,$E$88,$E$90:$E$91")
				.setValue("Error 80073cf9");
		ws.getRange("$G$1:$G$2,$I$1:$I$7,$H$8:$I$8,$A$93:$B$93,$E$93:$F$93").setValue(null);

		ws.getRange("$H$1:$H$7").setValue(new Object[][] { { "Constants" }, { "Formulas" }, { "String constants" },
				{ "Number constants" }, { "String formulas" }, { "Number formulas" }, { "Error formulas" } });

		ws.getRange("$A$2:$A$13").setValue("Test00001");
		ws.getRange("$A$14:$A$67").setValue("Test00153");
		ws.getRange("$A$68:$A$92").setValue("Test05789");
		ws.getRange("$E$3:$E$5,$E$9:$E$11,$E$25:$E$26,$E$37:$E$38,$E$57,$E$75:$E$76,$E$86:$E$87")
				.setValue("Runtime Error c0000005");
		ws.getRange(
				"$E$6,$E$13:$E$20,$E$28:$E$35,$E$41:$E$46,$E$52:$E$56,$E$61:$E$64,$E$72:$E$74,$E$77:$E$78,$E$82:$E$85,$E$89,$E$92")
				.setValue("Passed");
		ws.getRange("$E$8,$E$22:$E$24,$E$39,$E$49:$E$50,$E$58,$E$65:$E$69,$E$79").setValue("Deploy Error 80073cf9");

		ws.getRange("$D$2:$D$92").setFormulaR1C1("=\"X-Test-G\" & RC[-2] & \"-I\" & RC[-1]");
		ws.getRange("$B$3:$B$92").setFormulaR1C1("=IF(RC[-1]=R[-1]C[-1],R[-1]C,R[-1]C+1)");
		ws.getRange("$C$3:$C$92").setFormulaR1C1("=IF(RC[-2]=R[-1]C[-2],R[-1]C+1,1)");
		ws.getRange("$F$2:$F$92").setFormulaR1C1("=MID(RC[-1], SEARCH(\"Error \",RC[-1])+6,8)");

		Color constantBgColor = Color.FromArgb(0xFFDDEBF7);
		Color formulasBgColor = Color.FromArgb(0xFFF2F2F2);
		Color stringForeColor = Color.FromArgb(0xFF0000C0);
		Color errorForeColor = Color.GetDarkRed();

		IRange searchScope = ws.getRange("$A:$F");

		// Find constant cells and change background color
		IRange allConsts = searchScope.specialCells(SpecialCellType.Constants);
		allConsts.getInterior().setColor(constantBgColor);

		// Find formula cells and change background color
		IRange allFormulas = searchScope.specialCells(SpecialCellType.Formulas);
		allFormulas.getInterior().setColor(formulasBgColor);

		// Find text constant cells and change foreground color
		IRange textConsts = searchScope.specialCells(SpecialCellType.Constants, SpecialCellsValue.TextValues);
		textConsts.getFont().setColor(stringForeColor);

		// Find text formula cells and change foreground color
		IRange textFormulas = searchScope.specialCells(SpecialCellType.Formulas, SpecialCellsValue.TextValues);
		textFormulas.getFont().setColor(stringForeColor);

		// Find number constant cells and change font weight
		IRange numberConsts = searchScope.specialCells(SpecialCellType.Constants, SpecialCellsValue.Numbers);
		numberConsts.getFont().setBold(true);

		// Find number formula cells and change font weight
		IRange numberFormulas = searchScope.specialCells(SpecialCellType.Formulas, SpecialCellsValue.Numbers);
		numberFormulas.getFont().setBold(true);

		// Find error formula cells and change foreground color and font style
		IRange errorFormulas = searchScope.specialCells(SpecialCellType.Formulas, SpecialCellsValue.Errors);
		errorFormulas.getFont().setColor(errorForeColor);
		errorFormulas.getFont().setItalic(true);

		// Set sample cell styles
		ws.getRange("$H$1,$H$3,$H$4").getInterior().setColor(constantBgColor);
		ws.getRange("$H$2,$H$5:$H$7").getInterior().setColor(formulasBgColor);
		ws.getRange("$H$3,$H$5").getFont().setColor(stringForeColor);
		ws.getRange("$H$4,$H$6").getFont().setBold(true);
		ws.getRange("$H$7").getFont().setColor(errorForeColor);
		ws.getRange("$H$7").getFont().setItalic(true);

		ws.getUsedRange().getEntireColumn().autoFit();

	}

	@Override
	public boolean getIsNew() {
		return true;
	}
}
