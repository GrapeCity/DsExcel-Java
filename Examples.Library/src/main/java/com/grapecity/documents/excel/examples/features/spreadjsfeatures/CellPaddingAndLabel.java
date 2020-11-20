package com.grapecity.documents.excel.examples.features.spreadjsfeatures;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CellPaddingAndLabel extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("1:14").setRowHeight(40);
        worksheet.getRange("A:A").setColumnWidth(15);
        worksheet.getRange("B:B").setColumnWidth(10);
        worksheet.getRange("C:C").setColumnWidth(10);
        worksheet.getRange("D:D").setColumnWidth(10);
        worksheet.getRange("E:E").setColumnWidth(15);
        worksheet.getRange("F:F").setColumnWidth(25);
        worksheet.getRange("A1:F1").merge();
        worksheet.getRange("A2:F2").merge();
        worksheet.getRange("A3:F3").merge();
        worksheet.getRange("A4:F4").merge();
        worksheet.getRange("D5:F5").merge();
        worksheet.getRange("E6:F6").merge();
        worksheet.getRange("A7:D7").merge();
        worksheet.getRange("A8:D8").merge();
        worksheet.getRange("E8:F8").merge();
        worksheet.getRange("A9:F9").merge();
        worksheet.getRange("A10:D10").merge();
        worksheet.getRange("E10:F10").merge();
        worksheet.getRange("A11:E11").merge();
        worksheet.getRange("A12:C12").merge();
        worksheet.getRange("A13:F13").merge();
        worksheet.getRange("A14:F14").merge();
        worksheet.getRange("B5:C5").merge();
        worksheet.getRange("C6:D6").merge();

        worksheet.getRange("A1").setValue("Example Form");
        worksheet.getRange("A1").getFont().setBold(true);
        worksheet.getRange("A1").getFont().setSize(30);
        worksheet.getRange("A1").getFont().setName("Arial");

        worksheet.getRange("A2").setValue("Please open an account at");
        worksheet.getRange("A2").getFont().setBold(true);
        worksheet.getRange("A2").getFont().setSize(18);
        worksheet.getRange("A2").getFont().setName("Arial");

        worksheet.getRange("A3").setWatermark("BRANCH NAME");
        worksheet.getRange("A3").setCellPadding(new CellPadding(20, 0, 0, 0));
        worksheet.getRange("A3").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A3").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A3").getLabelOptions().setMargin(new Margin(2, 0, 0, 0));
        worksheet.getRange("A3").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A3").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A4").setValue("Personal Details");
        worksheet.getRange("A4").getFont().setBold(true);
        worksheet.getRange("A4").getFont().setSize(18);
        worksheet.getRange("A4").getFont().setName("Arial");
        worksheet.getRange("A4").setVerticalAlignment(VerticalAlignment.Bottom);

        worksheet.getRange("A5").setWatermark("MARITIAL STATUS");
        worksheet.getRange("A5").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A5").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A5").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A5").getLabelOptions().setMargin(new Margin(2, 0, 0, 0));
        worksheet.getRange("A5").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A5").getLabelOptions().getFont().setName("Arial");
        CheckBoxCellType cellType = new CheckBoxCellType();
        cellType.setTextFalse("Married");
        cellType.setTextTrue("Married");
        cellType.setTextAlign(CheckBoxAlign.Right);
        worksheet.getRange("A5").setCellType(cellType);

        worksheet.getRange("B5").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("B5").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("B5").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("B5").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("B5").getLabelOptions().getFont().setSize(10);
        worksheet.getRange("B5").getLabelOptions().getFont().setName("Arial");
        CheckBoxCellType cellType2 = new CheckBoxCellType();
        cellType2.setTextFalse("Single");
        cellType2.setTextTrue("Single");
        cellType2.setTextAlign(CheckBoxAlign.Right);
        worksheet.getRange("B5").setCellType(cellType2);

        worksheet.getRange("A6").setWatermark("EDUCATION");
        worksheet.getRange("A6").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A6").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A6").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A6").getLabelOptions().setMargin(new Margin(2, 0, 0, 0));
        worksheet.getRange("A6").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A6").getLabelOptions().getFont().setName("Arial");
        CheckBoxCellType cellType3 = new CheckBoxCellType();
        cellType3.setTextFalse("Under graduate");
        cellType3.setTextTrue("Under graduate");
        cellType3.setTextAlign(CheckBoxAlign.Right);
        worksheet.getRange("A6").setCellType(cellType3);


        worksheet.getRange("B6").setCellPadding(new CellPadding(15, 0, 0, 0));
        CheckBoxCellType cellType4 = new CheckBoxCellType();
        cellType4.setTextFalse("Graduate");
        cellType4.setTextTrue("Graduate");
        cellType4.setTextAlign(CheckBoxAlign.Right);
        worksheet.getRange("B6").setCellType(cellType4);

        worksheet.getRange("C6").setCellPadding(new CellPadding(15, 0, 0, 0));
        CheckBoxCellType cellType5 = new CheckBoxCellType();
        cellType5.setTextFalse("Others");
        cellType5.setTextTrue("Others");
        cellType5.setTextAlign(CheckBoxAlign.Right);
        worksheet.getRange("C6").setCellType(cellType5);

        worksheet.getRange("A7").setWatermark("E-MAIL");
        worksheet.getRange("A7").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A7").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A7").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A7").getLabelOptions().setMargin(new Margin(2, 0, 0, 0));
        worksheet.getRange("A7").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A7").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("E7").setWatermark("MOBILE NO.");
        worksheet.getRange("E7").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("E7").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("E7").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("E7").getLabelOptions().setMargin(new Margin(2, 0, 0, 0));
        worksheet.getRange("E7").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("E7").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("F7").setWatermark("EXISTING BANK ACCOUNT NO. (IF ANY)");
        worksheet.getRange("F7").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("F7").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("F7").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("F7").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("F7").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("F7").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("D5").setWatermark("FULL NAME");
        worksheet.getRange("D5").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("D5").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("D5").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("D5").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("D5").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("D5").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("E6").setWatermark("NATIONALITY");
        worksheet.getRange("E6").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("E6").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("E6").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("E6").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("E6").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("E6").getLabelOptions().getFont().setName("Arial");

        ComboBoxCellType cellType6 = new ComboBoxCellType();
        cellType6.setEditorValueType(EditorValueType.Text);
        ComboBoxCellItem item = new ComboBoxCellItem();
        item.setText("China");
        item.setValue("1");
        cellType6.getItems().add(item);

        ComboBoxCellItem item2 = new ComboBoxCellItem();
        item2.setText("United States");
        item2.setValue("1");
        cellType6.getItems().add(item2);

        ComboBoxCellItem item3 = new ComboBoxCellItem();
        item3.setText("Japan");
        item3.setValue("1");
        cellType6.getItems().add(item3);

        ComboBoxCellItem item4 = new ComboBoxCellItem();
        item4.setText("Germany");
        item4.setValue("1");
        cellType6.getItems().add(item4);

        ComboBoxCellItem item5 = new ComboBoxCellItem();
        item5.setText("France");
        item5.setValue("1");
        cellType6.getItems().add(item5);

        ComboBoxCellItem item6 = new ComboBoxCellItem();
        item6.setText("England");
        item6.setValue("1");
        cellType6.getItems().add(item6);

        worksheet.getRange("E6").setCellType(cellType6);

        worksheet.getRange("A8").setWatermark("IN CASE OF A MINOR PLEASE PROVIDE DETAILS");
        worksheet.getRange("A8").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A8").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A8").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A8").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("A8").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A8").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("E8").setWatermark("NAME OF FATHER/SPOUSE");
        worksheet.getRange("E8").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("E8").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("E8").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("E8").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("E8").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("E8").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A9").setValue("Residential address");
        worksheet.getRange("A9").setVerticalAlignment(VerticalAlignment.Bottom);
        worksheet.getRange("A9").getFont().setSize(18);
        worksheet.getRange("A9").getFont().setName("Arial");

        worksheet.getRange("A10").setWatermark("FLAT NO. AND BLDG. NAME");
        worksheet.getRange("A10").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A10").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A10").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A10").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("A10").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A10").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A11").setWatermark("ROAD NO./NAME");
        worksheet.getRange("A11").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A11").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A11").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A11").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("A11").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A11").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("E11").setWatermark("CITY");
        worksheet.getRange("E11").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("E11").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("E11").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("E11").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("E11").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("E11").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A12").setWatermark("TELEPHONE RESIDENCE");
        worksheet.getRange("A12").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A12").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A12").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A12").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("A12").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A12").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("D12").setWatermark("OFFICE");
        worksheet.getRange("D12").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("D12").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("D12").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("D12").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("D12").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("D12").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("E12").setWatermark("FAX");
        worksheet.getRange("E12").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("E12").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("E12").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("E12").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("E12").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("E12").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("F12").setWatermark("PIN CODE");
        worksheet.getRange("F12").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("F12").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("F12").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("F12").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("F12").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("F12").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A13").setValue("Referenced web site");
        worksheet.getRange("A13").setVerticalAlignment(VerticalAlignment.Bottom);
        worksheet.getRange("A13").getFont().setSize(18);
        worksheet.getRange("A13").getFont().setName("Arial");

        worksheet.getRange("A14").setWatermark("HELP MESSAGE");
        worksheet.getRange("A14").setCellPadding(new CellPadding(15, 0, 0, 0));
        worksheet.getRange("A14").getLabelOptions().setVisibility(LabelVisibility.visible);
        worksheet.getRange("A14").getLabelOptions().setForeColor(Color.FromArgb(51, 51, 51));
        worksheet.getRange("A14").getLabelOptions().setMargin(new Margin(2, 0, 0, 1));
        worksheet.getRange("A14").getLabelOptions().getFont().setSize(7);
        worksheet.getRange("A14").getLabelOptions().getFont().setName("Arial");

        worksheet.getRange("A14").getHyperlinks().add(worksheet.getRange("A14"),
                "http://www.grapecity.com/",
                null,
                "help",
                "Any question, please click here.");

        worksheet.getRange("A1:F15").getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("A1:F15").getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetBlack());

        worksheet.getRange("A1:F14").getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Thin);
        worksheet.getRange("A1:F14").getBorders().get(BordersIndex.InsideVertical).setColor(Color.GetBlack());
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
