package com.grapecity.documents.excel.examples.excelreporting;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.FontLanguageIndex;
import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IStyle;
import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.ITheme;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.TableStyleElementType;
import com.grapecity.documents.excel.Theme;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Themes;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.ConnectorType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ITextRange;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.Placement;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MovieList extends ExampleBase {

    @Override
    public void beforeExecute(Workbook workbook, String userAgents) {

        if (userAgents.toLowerCase().contains("macintosh")) {

            ITheme theme = new Theme("testTheme", Themes.GetOfficeTheme());
            theme.getThemeFontScheme().getMinor().get(FontLanguageIndex.Latin).setName("Trebuchet MS");
            workbook.setTheme(theme);
            IStyle style_Normal = workbook.getStyles().get("Normal");
            style_Normal.getFont().setThemeFont(ThemeFont.Minor);
        }
    }

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //------------------Set RowHeight & ColumnWidth----------------
        worksheet.setStandardHeight(43.5);
        worksheet.setStandardWidth(8.43);

        worksheet.getRange("1:1").setRowHeight(171);
        worksheet.getRange("2:2").setRowHeight(12.75);
        worksheet.getRange("3:3").setRowHeight(22.5);
        worksheet.getRange("4:7").setRowHeight(43.75);
        worksheet.getRange("A:A").setColumnWidth(2.887);
        worksheet.getRange("B:B").setColumnWidth(8.441);
        worksheet.getRange("C:C").setColumnWidth(12.777);
        worksheet.getRange("D:D").setColumnWidth(25.109);
        worksheet.getRange("E:E").setColumnWidth(12.109);
        worksheet.getRange("F:F").setColumnWidth(41.664);
        worksheet.getRange("G:G").setColumnWidth(18.555);
        worksheet.getRange("H:H").setColumnWidth(11);
        worksheet.getRange("I:I").setColumnWidth(13.664);
        worksheet.getRange("J:J").setColumnWidth(15.109);
        worksheet.getRange("K:K").setColumnWidth(38.887);
        worksheet.getRange("L:L").setColumnWidth(2.887);


        //------------------------Set Table Values-------------------
        ITable table = worksheet.getTables().add(worksheet.getRange("B3:K7"), true);
        worksheet.getRange("B3:K7").setValue(new Object[][]{
                {"NO.", "YEAR", "TITLE", "REVIEW", "STARRING ACTORS", "DIRECTOR", "GENRE", "RATING", "FORMAT", "COMMENTS"},
                {1, 1994, "Forrest Gump", "5 Stars", "Tom Hanks, Robin Wright, Gary Sinise", "Robert Zemeckis", "Drama", "PG-13", "DVD", "Based on the 1986 novel of the same name by Winston Groom"},
                {2, 1946, "It’s a Wonderful Life", "2 Stars", "James Stewart, Donna Reed, Lionel Barrymore ", "Frank Capra", "Drama", "G", "VHS", "Colorized version"},
                {3, 1988, "Big", "4 Stars", "Tom Hanks, Elizabeth Perkins, Robert Loggia ", "Penny Marshall", "Comedy", "PG", "DVD", ""},
                {4, 1954, "Rear Window", "3 Stars", "James Stewart, Grace Kelly, Wendell Corey ", "Alfred Hitchcock", "Suspense", "PG", "Blu-ray", ""},
        });


        //-----------------------Set Table style--------------------
        ITableStyle tableStyle = workbook.getTableStyles().add("Movie List");
        workbook.setDefaultTableStyle("Movie List");

        tableStyle.getTableStyleElements().get(TableStyleElementType.WholeTable).getInterior().setColor(Color.GetWhite());

        tableStyle.getTableStyleElements().get(TableStyleElementType.FirstRowStripe).getInterior().setColor(Color.FromArgb(38, 38, 38));

        tableStyle.getTableStyleElements().get(TableStyleElementType.SecondRowStripe).getInterior().setColor(Color.GetBlack());

        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getFont().setColor(Color.GetBlack());
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().setColor(Color.FromArgb(38, 38, 38));
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getInterior().setColor(Color.FromArgb(68, 217, 255));
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Thick);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeLeft).setLineStyle(BorderLineStyle.None);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeRight).setLineStyle(BorderLineStyle.None);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.None);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.None);
        tableStyle.getTableStyleElements().get(TableStyleElementType.HeaderRow).getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.None);


        //--------------------------------Set Named Styles---------------------
        IStyle movieListBorderStyle = workbook.getStyles().add("Movie list border");
        movieListBorderStyle.setIncludeNumber(true);
        movieListBorderStyle.setIncludeAlignment(true);
        movieListBorderStyle.setVerticalAlignment(VerticalAlignment.Center);
        movieListBorderStyle.setWrapText(true);
        movieListBorderStyle.setIncludeFont(true);
        movieListBorderStyle.getFont().setName("Helvetica");
        movieListBorderStyle.getFont().setSize(11);
        movieListBorderStyle.getFont().setColor(Color.GetWhite());
        movieListBorderStyle.setIncludeBorder(true);
        movieListBorderStyle.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        movieListBorderStyle.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(38, 38, 38));
        movieListBorderStyle.setIncludePatterns(true);
        movieListBorderStyle.getInterior().setColor(Color.FromArgb(238, 219, 78));

        IStyle nOStyle = workbook.getStyles().add("NO.");
        nOStyle.setIncludeNumber(true);
        nOStyle.setIncludeAlignment(true);
        nOStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        nOStyle.setVerticalAlignment(VerticalAlignment.Center);
        nOStyle.setIncludeFont(true);
        nOStyle.getFont().setName("Helvetica");
        nOStyle.getFont().setSize(11);
        nOStyle.getFont().setColor(Color.GetWhite());
        nOStyle.setIncludeBorder(true);
        nOStyle.setIncludePatterns(true);
        nOStyle.getInterior().setColor(Color.FromArgb(38, 38, 38));

        IStyle reviewStyle = workbook.getStyles().add("Review");
        reviewStyle.setIncludeNumber(true);
        reviewStyle.setIncludeAlignment(true);
        reviewStyle.setVerticalAlignment(VerticalAlignment.Center);
        reviewStyle.setIncludeFont(true);
        reviewStyle.getFont().setName("Helvetica");
        reviewStyle.getFont().setSize(11);
        reviewStyle.getFont().setColor(Color.GetWhite());
        reviewStyle.setIncludeBorder(true);
        reviewStyle.setIncludePatterns(true);
        reviewStyle.getInterior().setColor(Color.FromArgb(38, 38, 38));

        IStyle yearStyle = workbook.getStyles().add("Year");
        yearStyle.setIncludeNumber(true);
        yearStyle.setIncludeAlignment(true);
        yearStyle.setHorizontalAlignment(HorizontalAlignment.Left);
        yearStyle.setVerticalAlignment(VerticalAlignment.Center);
        yearStyle.setIncludeFont(true);
        yearStyle.getFont().setName("Helvetica");
        yearStyle.getFont().setSize(11);
        yearStyle.getFont().setColor(Color.GetWhite());
        yearStyle.setIncludeBorder(true);
        yearStyle.setIncludePatterns(true);
        yearStyle.getInterior().setColor(Color.FromArgb(38, 38, 38));

        IStyle heading1Style = workbook.getStyles().get("Heading 1");
        heading1Style.setIncludeAlignment(true);
        heading1Style.setVerticalAlignment(VerticalAlignment.Bottom);
        heading1Style.setIncludeBorder(true);
        heading1Style.getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Thick);
        heading1Style.getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(68, 217, 255));
        heading1Style.setIncludeFont(true);
        heading1Style.getFont().setName("Helvetica");
        heading1Style.getFont().setBold(false);
        heading1Style.getFont().setSize(12);
        heading1Style.getFont().setColor(Color.GetBlack());

        IStyle normalStyle = workbook.getStyles().get("Normal");
        normalStyle.setIncludeNumber(true);
        normalStyle.setIncludeAlignment(true);
        normalStyle.setVerticalAlignment(VerticalAlignment.Center);
        normalStyle.setWrapText(true);
        normalStyle.setIncludeFont(true);
        normalStyle.getFont().setName("Helvetica");
        normalStyle.getFont().setSize(11);
        normalStyle.getFont().setColor(Color.GetWhite());
        normalStyle.setIncludePatterns(true);
        normalStyle.getInterior().setColor(Color.FromArgb(38, 38, 38));


        //-----------------------------Use NamedStyle--------------------------
        worksheet.getSheetView().setDisplayGridlines(false);
        worksheet.setTabColor(Color.FromArgb(38, 38, 38));
        table.setTableStyle(tableStyle);

        worksheet.getRange("A2:L2").setStyle(movieListBorderStyle);
        worksheet.getRange("B3:K3").setStyle(heading1Style);
        worksheet.getRange("B4:B7").setStyle(nOStyle);
        worksheet.getRange("C4:C7").setStyle(yearStyle);
        worksheet.getRange("E4:E7").setStyle(reviewStyle);
        worksheet.getRange("F4:F7").setIndentLevel(1);
        worksheet.getRange("F4:F7").setHorizontalAlignment(HorizontalAlignment.Left);


        //-----------------------------Add Shapes------------------------------
        //Movie picture
        InputStream stream = this.getResourceStream("movie.png");

        //Movie list picture
        InputStream stream2 = this.getResourceStream("list.png");
        try {
            IShape pictureShape = worksheet.getShapes().addPicture(stream, ImageType.PNG, 0, 1, worksheet.getRange("A:L").getWidth(), worksheet.getRange("1:1").getHeight() - 1.5);
            pictureShape.setPlacement(Placement.Move);

            IShape pictureShape2 = worksheet.getShapes().addPicture(stream2, ImageType.PNG, 1, 0.8, 325.572, 85.51);
            pictureShape2.setPlacement(Placement.Move);

        } catch (IOException e) {
            e.printStackTrace();
        }

        //Rounded Rectangular Callout 7
        IShape roundedRectangular = worksheet.getShapes().addShape(AutoShapeType.RoundedRectangularCallout, 437.5, 22.75, 342, 143);
        roundedRectangular.setName("Rounded Rectangular Callout 7");
        roundedRectangular.setPlacement(Placement.Move);
        roundedRectangular.getTextFrame().getTextRange().getFont().setName("Helvetica");
        roundedRectangular.getTextFrame().getTextRange().getFont().getColor().setRGB(Color.FromArgb(38, 38, 38));

        roundedRectangular.getFill().solid();
        roundedRectangular.getFill().getColor().setRGB(Color.FromArgb(68, 217, 255));
        roundedRectangular.getFill().setTransparency(0);
        roundedRectangular.getLine().solid();
        roundedRectangular.getLine().getColor().setRGB(Color.FromArgb(0, 129, 162));
        roundedRectangular.getLine().setWeight(2);
        roundedRectangular.getLine().setTransparency(0);

        ITextRange roundedRectangular_p0 = roundedRectangular.getTextFrame().getTextRange().getParagraphs().get(0);
        roundedRectangular_p0.getRuns().getFont().setBold(true);
        roundedRectangular_p0.getRuns().add("TABLE");
        roundedRectangular_p0.getRuns().add(" TIP");

        roundedRectangular.getTextFrame().getTextRange().getParagraphs().add("");

        ITextRange roundedRectangular_p2 = roundedRectangular.getTextFrame().getTextRange().getParagraphs().add();
        roundedRectangular_p2.getRuns().add("Use the drop down arrows in the table headings to quickly filter your movie list. " +
                "For multiple entry fields, such as Starring Actors,  select the drop down arrow next to the field and enter text in the Search box. " +
                "For example, type Tom Hanks or James Stewart, and then select OK.");

        roundedRectangular.getTextFrame().getTextRange().getParagraphs().add("");

        ITextRange roundedRectangular_p4 = roundedRectangular.getTextFrame().getTextRange().getParagraphs().add();
        roundedRectangular_p4.getRuns().add("To delete this note, click the edge to select it and then press ");
        roundedRectangular_p4.getRuns().add("Delete");
        roundedRectangular_p4.getRuns().add(".");
        roundedRectangular_p4.getRuns().get(2).getFont().setBold(true);

        roundedRectangular.getTextFrame().getTextRange().getParagraphs().add("");

        //Add Stright Line Shape
        IShape lineShape = worksheet.getShapes().addConnector(ConnectorType.Straight, 455.228f, 57.35f, 756.228f, 57.35f);
        lineShape.getLine().solid();
        lineShape.getLine().setWeight(3);
        lineShape.getLine().getColor().setRGB(Color.FromArgb(38, 38, 38));
        lineShape.getLine().setDashStyle(LineDashStyle.SysDot);

    }


    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public String[] getResources() {
        return new String[]{"movie.png", "list.png"};
    }
}
