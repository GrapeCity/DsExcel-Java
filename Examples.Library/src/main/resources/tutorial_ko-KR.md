
# 스프레드시트 API - Documents for Excel - Java Edition 시작하기

이 자습서는 GrapeCity Documents for Excel의 기능에 대한 이해를 돕기위해 따라하실수 있는 실제적인 시나리오를 만들었습니다. 자습서를 마치시면 기초적인 예산용 Excel 파일이 완성됩니다.

## 준비

1. JDK 6 이상 버전이 설치되어 있어야 합니다.

2. Intellij IDEA 또는 Eclipse와 같은 선호하는 Java IDE로 Java 콘솔 응용 프로그램을 만듭니다.

3. 다음과 같이 GcExcel 라이브러리의 종속성을 추가합니다.
> **Intellij 또는 Eclipse 콘솔 응용 프로그램** 
> - [Maven](https://search.maven.org/artifact/com.grapecity.documents/gcexcel/) 또는 [Github](https://github.com/GrapeCity/GcExcel-Java)에서 gcexcel jar 패키지 다운로드 
> - gcexcel-4.0.0.jar을 프로젝트 라이브러리 폴더에 복사하고 해당 라이브러리로 추가
>
> **Gradle 프로젝트** 
> - build.gradle을 열고 해당 항목에 아래 스크립트 추가
> ```xml
> compile("com.grapecity.documents:gcexcel:4.0.0")
> ```
>
> **Maven 프로젝트** 
> - pom.xml을 열고 해당 노드에 아래 xml 요소 추가
> ```xml
> <dependency>
>    <groupId>com.grapecity.documents</groupId>
>    <artifactId>gcexcel</artifactId>
>    <version>4.0.0</version>
> </dependency>
> ```

## 네임스페이스 추가

Main.java를 열고 아래 네임스페이스를 가져옵니다.

- Java
```java
import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.*;
```

- Kotlin
```kotlin
import com.grapecity.documents.excel.*
import com.grapecity.documents.excel.drawing.*
```

## 통합 문서 만들기

첫번째 단계는 GrapeCity Documents for Excel API를 사용하여 Excel 파일을 만드는 것입니다.

- Java
```java
Workbook workbook = new Workbook();
IWorksheet worksheet = workbook.getWorksheets().get(0);
```

- Kotlin
```kotlin
var workbook = new Workbook()
var worksheet = workbook.getWorksheets().get(0)
```

## 데이터 초기화

그리고 GcExcel에서 데이터를 초기화하려면 2차원 배열을 준비하고 이를 워크시트 범위의 값에 할당합니다.

- Java
```java
worksheet.getRange("B3:C7").setValue(new Object[][]{
    {"ITEM", "AMOUNT"},
    {"Income 1", 2500},
    {"Income 2", 1000},
    {"Income 3", 250},
    {"Other", 250},
});
worksheet.getRange("B10:C23").setValue(new Object[][]{
    {"ITEM", "AMOUNT"},
    {"Rent/mortgage", 800},
    {"Electric", 120},
    {"Gas", 50},
    {"Cell phone", 45},
    {"Groceries", 500},
    {"Car payment", 273},
    {"Auto expenses", 120},
    {"Student loans", 50},
    {"Credit cards", 100},
    {"Auto Insurance", 78},
    {"Personal care", 50},
    {"Entertainment", 100},
    {"Miscellaneous", 50},
});

worksheet.getRange("B2:C2").merge();
worksheet.getRange("B2").setValue("MONTHLY INCOME");
worksheet.getRange("B9:C9").merge();
worksheet.getRange("B9").setValue("MONTHLY EXPENSES");
worksheet.getRange("E2:G2").merge();
worksheet.getRange("E2").setValue("PERCENTAGE OF INCOME SPENT");
worksheet.getRange("E5:G5").merge();
worksheet.getRange("E5").setValue("SUMMARY");
worksheet.getRange("E3:F3").merge();
worksheet.getRange("E9").setValue("BALANCE");
worksheet.getRange("E6").setValue("Total Monthly Income");
worksheet.getRange("E7").setValue("Total Monthly Expenses");
```

- Kotkin
```kotlin
worksheet.getRange("B3:C7").setValue(
    arrayOf<Array<Any>>(
            arrayOf<Any>("ITEM", "AMOUNT"),
            arrayOf<Any>("Income 1", 2500),
            arrayOf<Any>("Income 2", 1000),
            arrayOf<Any>("Income 3", 250),
            arrayOf<Any>("Other", 250)
    )
)
worksheet.getRange("B10:C23").setValue(
    arrayOf<Array<Any>>(
            arrayOf<Any>("ITEM", "AMOUNT"),
            arrayOf<Any>("Rent/mortgage", 800),
            arrayOf<Any>("Electric", 120),
            arrayOf<Any>("Gas", 50),
            arrayOf<Any>("Cell phone", 45),
            arrayOf<Any>("Groceries", 500),
            arrayOf<Any>("Car payment", 273),
            arrayOf<Any>("Auto expenses", 120),
            arrayOf<Any>("Student loans", 50),
            arrayOf<Any>("Credit cards", 100),
            arrayOf<Any>("Auto Insurance", 78),
            arrayOf<Any>("Personal care", 50),
            arrayOf<Any>("Entertainment", 100),
            arrayOf<Any>("Miscellaneous", 50)
    )
)
worksheet.getRange("B2:C2").merge()
worksheet.getRange("B2").setValue("MONTHLY INCOME")
worksheet.getRange("B9:C9").merge()
worksheet.getRange("B9").setValue("MONTHLY EXPENSES")
worksheet.getRange("E2:G2").merge()
worksheet.getRange("E2").setValue("PERCENTAGE OF INCOME SPENT")
worksheet.getRange("E5:G5").merge()
worksheet.getRange("E5").setValue("SUMMARY")
worksheet.getRange("E3:F3").merge()
worksheet.getRange("E9").setValue("BALANCE")
worksheet.getRange("E6").setValue("Total Monthly Income")
worksheet.getRange("E7").setValue("Total Monthly Expenses")
```

## 행 높이 및 열 너비 설정

행 높이와 열 너비를 사용자 정의하고 레이아웃 및 데이터 프레젠테이션을 설정합니다. "StandardHeight" 및 "StandardWidth"를 사용하여 워크시트의 기본 행 높이와 열 너비를 설정할 수 있습니다.

- Java
```java
worksheet.setStandardHeight(26.25);
worksheet.setStandardWidth(8.43);

worksheet.getRange("2:24").setRowHeight(27);
worksheet.getRange("A:A").setColumnWidth(2.855);
worksheet.getRange("B:B").setColumnWidth(33.285);
worksheet.getRange("C:C").setColumnWidth(25.57);
worksheet.getRange("D:D").setColumnWidth(1);
worksheet.getRange("E:F").setColumnWidth(25.57);
worksheet.getRange("G:G").setColumnWidth(14.285);
```

- Kotlin
```kotlin
worksheet.setStandardHeight(26.25)
worksheet.setStandardWidth(8.43)
worksheet.getRange("2:24").setRowHeight(27.0)
worksheet.getRange("A:A").setColumnWidth(2.855)
worksheet.getRange("B:B").setColumnWidth(33.285)
worksheet.getRange("C:C").setColumnWidth(25.57)
worksheet.getRange("D:D").setColumnWidth(1.0)
worksheet.getRange("E:F").setColumnWidth(25.57)
worksheet.getRange("G:G").setColumnWidth(14.285)
```

## 표 만들기

두 개의 표("수입" 및 "비용")를 추가하고 각 표에 기본 표 스타일을 적용합니다.

- Java
```java
ITable incomeTable = worksheet.getTables().add(worksheet.getRange("B3:C7"), true);
incomeTable.setName("tblIncome");
incomeTable.setTableStyle(workbook.getTableStyles().get("TableStyleMedium4"));

ITable expensesTable = worksheet.getTables().add(worksheet.getRange("B10:C23"), true);
expensesTable.setName("tblExpenses");
expensesTable.setTableStyle(workbook.getTableStyles().get("TableStyleMedium4"));
```

- Kotlin
```kotlin
val incomeTable = worksheet.getTables().add(worksheet.getRange("B3:C7"), true)
incomeTable.setName("tblIncome")
incomeTable.setTableStyle(workbook.getTableStyles().get("TableStyleMedium4"))
val expensesTable = worksheet.getTables().add(worksheet.getRange("B10:C23"), true)
expensesTable.setName("tblExpenses")
expensesTable.setTableStyle(workbook.getTableStyles().get("TableStyleMedium4"))
```

## 함수 설정

사용자 정의 이름 2개를 만들어 월수입과 비용을 할당하고 월수입 합계, 월 비용 합계, 수입 대비 지출 백분율, 잔액을 계산하는 함수를 추가합니다.

- Java
```java
worksheet.getNames().add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])");
worksheet.getNames().add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])");

worksheet.getRange("E3").setFormula("=TotalMonthlyExpenses");
worksheet.getRange("G3").setFormula("=TotalMonthlyExpenses/TotalMonthlyIncome");
worksheet.getRange("G6").setFormula("=TotalMonthlyIncome");
worksheet.getRange("G7").setFormula("=TotalMonthlyExpenses");
worksheet.getRange("G9").setFormula("=TotalMonthlyIncome-TotalMonthlyExpenses");
```

- Kotlin
```kotlin
worksheet.getNames().add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])")
worksheet.getNames().add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])")
worksheet.getRange("E3").setFormula("=TotalMonthlyExpenses")
worksheet.getRange("G3").setFormula("=TotalMonthlyExpenses/TotalMonthlyIncome")
worksheet.getRange("G6").setFormula("=TotalMonthlyIncome")
worksheet.getRange("G7").setFormula("=TotalMonthlyExpenses")
worksheet.getRange("G9").setFormula("=TotalMonthlyIncome-TotalMonthlyExpenses")
```


## 스타일 설정

범위의 스타일을 변경하는 방법은 두 가지입니다. 
- 기본 제공 스타일 또는 사용자 정의 스타일을 이름으로 적용 
- 각 요소에 개별 스타일 설정

"통화", "제목 1", "백분율" 기본 제공 스타일을 정의하여 셀 범위에 적용합니다. 다른 범위에 적용할 개별 스타일 요소도 정의합니다.

- Java
```java
IStyle currencyStyle = workbook.getStyles().get("Currency");
currencyStyle.setIncludeAlignment(true);
currencyStyle.setHorizontalAlignment(HorizontalAlignment.Left);
currencyStyle.setVerticalAlignment(VerticalAlignment.Bottom);
currencyStyle.setNumberFormat("$#,##0.00");

IStyle heading1Style = workbook.getStyles().get("Heading 1");
heading1Style.setIncludeAlignment(true);
heading1Style.setHorizontalAlignment(HorizontalAlignment.Center);
heading1Style.setVerticalAlignment(VerticalAlignment.Center);
heading1Style.getFont().setName("Century Gothic");
heading1Style.getFont().setBold(true);
heading1Style.getFont().setSize(11);
heading1Style.getFont().setColor(Color.GetWhite());
heading1Style.setIncludeBorder(false);
heading1Style.setIncludePatterns(true);
heading1Style.getInterior().setColor(Color.FromArgb(32, 61, 64));

IStyle percentStyle = workbook.getStyles().get("Percent");
percentStyle.setIncludeAlignment(true);
percentStyle.setHorizontalAlignment(HorizontalAlignment.Center);
percentStyle.setIncludeFont(true);
percentStyle.getFont().setColor(Color.FromArgb(32, 61, 64));
percentStyle.getFont().setName("Century Gothic");
percentStyle.getFont().setBold(true);
percentStyle.getFont().setSize(14);

worksheet.getSheetView().setDisplayGridlines(false);
worksheet.getRange("C4:C7, C11:C23, G6:G7, G9").setStyle(currencyStyle);
worksheet.getRange("B2, B9, E2, E5").setStyle(heading1Style);
worksheet.getRange("G3").setStyle(percentStyle);

worksheet.getRange("E6:G6").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Medium);
worksheet.getRange("E6:G6").getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(32, 61, 64));
worksheet.getRange("E7:G7").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Medium);
worksheet.getRange("E7:G7").getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(32, 61, 64));

worksheet.getRange("E9:G9").getInterior().setColor(Color.FromArgb(32, 61, 64));
worksheet.getRange("E9:G9").setHorizontalAlignment(HorizontalAlignment.Left);
worksheet.getRange("E9:G9").setVerticalAlignment(VerticalAlignment.Center);
worksheet.getRange("E9:G9").getFont().setName("Century Gothic");
worksheet.getRange("E9:G9").getFont().setBold(true);
worksheet.getRange("E9:G9").getFont().setSize(11);
worksheet.getRange("E9:G9").getFont().setColor(Color.GetWhite());
worksheet.getRange("E3:F3").getBorders().setColor(Color.FromArgb(32, 61, 64));
```

- Kotlin
```kotlin
val currencyStyle = workbook.getStyles().get("Currency")
currencyStyle.setIncludeAlignment(true)
currencyStyle.setHorizontalAlignment(HorizontalAlignment.Left)
currencyStyle.setVerticalAlignment(VerticalAlignment.Bottom)
currencyStyle.setNumberFormat("$#,##0.00")

val heading1Style = workbook.getStyles().get("Heading 1")
heading1Style.setIncludeAlignment(true)
heading1Style.setHorizontalAlignment(HorizontalAlignment.Center)
heading1Style.setVerticalAlignment(VerticalAlignment.Center)
heading1Style.getFont().setName("Century Gothic")
heading1Style.getFont().setBold(true)
heading1Style.getFont().setSize(11.0)
heading1Style.getFont().setColor(Color.GetWhite())
heading1Style.setIncludeBorder(false)
heading1Style.setIncludePatterns(true)
heading1Style.getInterior().setColor(Color.FromArgb(32, 61, 64))

val percentStyle = workbook.getStyles().get("Percent")
percentStyle.setIncludeAlignment(true)
percentStyle.setHorizontalAlignment(HorizontalAlignment.Center)
percentStyle.setIncludeFont(true)
percentStyle.getFont().setColor(Color.FromArgb(32, 61, 64))
percentStyle.getFont().setName("Century Gothic")
percentStyle.getFont().setBold(true)
percentStyle.getFont().setSize(14.0)

worksheet.getSheetView().setDisplayGridlines(false)

worksheet.getRange("C4:C7, C11:C23, G6:G7, G9").setStyle(currencyStyle)
worksheet.getRange("B2, B9, E2, E5").setStyle(heading1Style)
worksheet.getRange("G3").setStyle(percentStyle)
worksheet.getRange("E6:G6").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Medium)
worksheet.getRange("E6:G6").getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(32, 61, 64))
worksheet.getRange("E7:G7").getBorders().get(BordersIndex.EdgeBottom).setLineStyle(BorderLineStyle.Medium)
worksheet.getRange("E7:G7").getBorders().get(BordersIndex.EdgeBottom).setColor(Color.FromArgb(32, 61, 64))
worksheet.getRange("E9:G9").getInterior().setColor(Color.FromArgb(32, 61, 64))
worksheet.getRange("E9:G9").setHorizontalAlignment(HorizontalAlignment.Left)
worksheet.getRange("E9:G9").setVerticalAlignment(VerticalAlignment.Center)
worksheet.getRange("E9:G9").getFont().setName("Century Gothic")
worksheet.getRange("E9:G9").getFont().setBold(true)
worksheet.getRange("E9:G9").getFont().setSize(11.0)
worksheet.getRange("E9:G9").getFont().setColor(Color.GetWhite())
worksheet.getRange("E3:F3").getBorders().setColor(Color.FromArgb(32, 61, 64))
```

## 조건부 서식 추가

GrapeCity Documents for Excel은 모든 유형의 조건부 서식을 지원합니다. 그라데이션 데이터 막대 서식으로 수입 대비 지출 백분율을 표시할 수 있습니다. 이 규칙은 값을 표시하지 않고 데이터 막대만을 표시합니다.

- Java
```java
IDataBar dataBar = worksheet.getRange("E3").getFormatConditions().addDatabar();
dataBar.getMinPoint().setType(ConditionValueTypes.Number);
dataBar.getMinPoint().setValue(1);
dataBar.getMaxPoint().setType(ConditionValueTypes.Number);
dataBar.getMaxPoint().setValue("=TotalMonthlyIncome");
dataBar.setBarFillType(DataBarFillType.Gradient);
dataBar.getBarColor().setColor(Color.GetRed());
dataBar.setShowValue(false);
```

Kotlin
```kotlin
val dataBar = worksheet.getRange("E3").getFormatConditions().addDatabar()
dataBar.getMinPoint().setType(ConditionValueTypes.Number)
dataBar.getMinPoint().setValue(1)
dataBar.getMaxPoint().setType(ConditionValueTypes.Number)
dataBar.getMaxPoint().setValue("=TotalMonthlyIncome")
dataBar.setBarFillType(DataBarFillType.Gradient)
dataBar.getBarColor().setColor(Color.GetRed())
dataBar.setShowValue(false)
```

## 차트 추가 

세로 막대형 차트를 만들어 수입과 비용의 차이를 비교합니다. 계열 겹침과 간격 너비를 변경하고 차트 영역, 축 선, 눈금 레이블, 데이터 요소 등 몇 가지 차트 요소의 서식을 사용자 정의하여 레이아웃을 만들수 있습니다.

- Java
```java
IShape shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 339, 247, 316.5, 346);
shape.getChart().getChartArea().getFormat().getLine().setTransparency(1);
shape.getChart().getColumnGroups().get(0).setOverlap(0);
shape.getChart().getColumnGroups().get(0).setGapWidth(37);

IAxis category_axis = shape.getChart().getAxes().item(AxisType.Category);
category_axis.getFormat().getLine().getColor().setRGB(Color.GetBlack());
category_axis.getTickLabels().getFont().setSize(11);
category_axis.getTickLabels().getFont().getColor().setRGB(Color.GetBlack());

IAxis series_axis = shape.getChart().getAxes().item(AxisType.Value);
series_axis.getFormat().getLine().setWeight(1);
series_axis.getFormat().getLine().getColor().setRGB(Color.GetBlack());
series_axis.getTickLabels().setNumberFormat("$###0");
series_axis.getTickLabels().getFont().setSize(11);
series_axis.getTickLabels().getFont().getColor().setRGB(Color.GetBlack());

ISeries chartSeries = shape.getChart().getSeriesCollection().newSeries();
chartSeries.setFormula("=SERIES(\"Simple Budget\",{\"Income\",\"Expenses\"},'Sheet1'!$G$6:$G$7,1)");
chartSeries.getPoints().get(0).getFormat().getFill().getColor().setRGB(Color.FromArgb(176, 21, 19));
chartSeries.getPoints().get(1).getFormat().getFill().getColor().setRGB(Color.FromArgb(234, 99, 18));
chartSeries.getDataLabels().getFont().setSize(11);
chartSeries.getDataLabels().getFont().getColor().setRGB(Color.GetBlack());
chartSeries.getDataLabels().setShowValue(true);
chartSeries.getDataLabels().setPosition(DataLabelPosition.OutsideEnd);
```

- Kotlin
```kotlin
val shape = worksheet.getShapes().addChart(ChartType.ColumnClustered, 339.0, 247.0, 3
shape.getChart().getChartArea().getFormat().getLine().setTransparency(1.0)
shape.getChart().getColumnGroups().get(0).setOverlap(0)
shape.getChart().getColumnGroups().get(0).setGapWidth(37)

val category_axis = shape.getChart().getAxes().item(AxisType.Category)
category_axis.getFormat().getLine().getColor().setRGB(Color.GetBlack())
category_axis.getTickLabels().getFont().setSize(11.0)
category_axis.getTickLabels().getFont().getColor().setRGB(Color.GetBlack())

val series_axis = shape.getChart().getAxes().item(AxisType.Value)
series_axis.getFormat().getLine().setWeight(1.0)
series_axis.getFormat().getLine().getColor().setRGB(Color.GetBlack())
series_axis.getTickLabels().setNumberFormat("$###0")
series_axis.getTickLabels().getFont().setSize(11.0)
series_axis.getTickLabels().getFont().getColor().setRGB(Color.GetBlack())

val chartSeries = shape.getChart().getSeriesCollection().newSeries()
chartSeries.setFormula("=SERIES(\"Simple Budget\",{\"Income\",\"Expenses\"},'Sheet1'!
chartSeries.getPoints().get(0).getFormat().getFill().getColor().setRGB(Color.FromArgb
chartSeries.getPoints().get(1).getFormat().getFill().getColor().setRGB(Color.FromArgb
chartSeries.getDataLabels().getFont().setSize(11.0)
chartSeries.getDataLabels().getFont().getColor().setRGB(Color.GetBlack())
chartSeries.getDataLabels().setShowValue(true)
chartSeries.getDataLabels().setPosition(DataLabelPosition.OutsideEnd)
```

## Excel로 저장

"SimpleBudget.xlsx"라는 이름으로 Excel 파일로 저장합니다.

- Java
```java
workbook.save("SimpleBudget.xlsx");
```

- Kotlin
```kotlin
workbook.save("SimpleBudget.xlsx")
```

최종 결과물은 아래의 경로에서 [SimpleBudget.xlsx](api/examples/xlsx/tutorial?fileName=SimpleBudget)확인하여 볼수 있습니다.[듀토리얼 소스 프로젝트](resources/gcexcel-tutorial.zip)를 다운로드하고 코드를 직접 실행하고 싶으신 경우에는 경우 먼저 JAVA가 컴퓨터에 설치되어 있어야 합니다.