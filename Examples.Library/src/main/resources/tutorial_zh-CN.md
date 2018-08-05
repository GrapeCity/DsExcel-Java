# 介绍

在本教程中, 我们将会通过一个真实的使用场景，来让大家对Sprea表格组件能够做什么有一个基础的了解，并且在教程的最后一步，你将会得到一个用Spread表格组件生成的Excel文件，这个文件可以用来分析你的月度收入和支出。

## 准备

1. 安装[.NET Core](https://www.microsoft.com/net/core). 虽然在本教程中我们使用.NET Core, 但你可以用类似的方式在.NET Framework和Mono中去使用Spread表格组件.

2. 在**Visual Studio 2017**中创建一个.NET Core的控制台程序或者更简单的方式是使用下面的**dotnet CLI**.
> ```csharp
> dotnet new console
> ```

3. 安装**Spread表格组件** NuGet包:
> **Visual Studio**
> - 鼠标右键单击工程文件, 选择 "管理NuGet程序包..."。
> - 选择**nuget.org**作为程序包源, 然后搜索"GrapeCity.Documents.Excel"，最后点击安装。
>
> **dotnet CLI** 
> - 在当前工程文件所在的文件路径下打开一个命令行窗口。
> - 执行下面的命令行程序:
> ```csharp
> dotnet add package GrapeCity.Documents.Excel
> ```

## 引入名字空间

打开**Program.cs**文件，然后引入以下两个名字空间

```csharp
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing; 
```

## 创建工作簿

通过Spread表格组件来生成一个Excel文件的第一步就是先创建一个工作簿实例。

```csharp
Workbook workbook = new Workbook();
IWorksheet worksheet = workbook.Worksheets[0];
```

## 初始化数据

在给**Spread表格组件**设置大量数据时, 需要先准备一个填充好的二维数组，然后将这个二维数组赋值给需要设置数据的区域。

```csharp
worksheet.Range["B3:C7"].Value = new object[,]
{
    { "ITEM", "AMOUNT" },
    { "Income 1", 2500 },
    { "Income 2", 1000 },
    { "Income 3", 250 },
    { "Other", 250 },
};

worksheet.Range["B10:C23"].Value = new object[,]
{
    { "ITEM", "AMOUNT" },
    { "Rent/mortgage", 800 },
    { "Electric", 120 },
    { "Gas", 50 },
    { "Cell phone", 45 },
    { "Groceries", 500 },
    { "Car payment", 273 },
    { "Auto expenses", 120 },
    { "Student loans", 50 },
    { "Credit cards", 100 },
    { "Auto Insurance", 78 },
    { "Personal care", 50 },
    { "Entertainment", 100 },
    { "Miscellaneous", 50 },
};

worksheet.Range["B2:C2"].Merge();
worksheet.Range["B2"].Value = "MONTHLY INCOME";
worksheet.Range["B9:C9"].Merge();
worksheet.Range["B9"].Value = "MONTHLY EXPENSES";
worksheet.Range["E2:G2"].Merge();
worksheet.Range["E2"].Value = "PERCENTAGE OF INCOME SPENT";
worksheet.Range["E5:G5"].Merge();
worksheet.Range["E5"].Value = "SUMMARY";
worksheet.Range["E3:F3"].Merge();
worksheet.Range["E9"].Value = "BALANCE";
worksheet.Range["E6"].Value = "Total Monthly Income";
worksheet.Range["E7"].Value = "Total Monthly Expenses";
```

## 设置行高和列宽

有时需要设置合适的行高和列宽来让数据展示得更好，可以使用Worksheet的StandardHeight和StandardWidth属性来设置默认的行高和列宽。

```csharp
worksheet.StandardHeight = 26.25;
worksheet.StandardWidth = 8.43;

worksheet.Range["2:24"].RowHeight = 27;
worksheet.Range["A:A"].ColumnWidth = 2.855;
worksheet.Range["B:B"].ColumnWidth = 33.285;
worksheet.Range["C:C"].ColumnWidth = 25.57;
worksheet.Range["D:D"].ColumnWidth = 1;
worksheet.Range["E:F"].ColumnWidth = 25.57;
worksheet.Range["G:G"].ColumnWidth = 14.285;
```

## 创建表格

创建两个表格: "Income"和"Expenses"，并分别给这两个表格设置内置的表格样式。

```csharp
ITable incomeTable = worksheet.Tables.Add(worksheet.Range["B3:C7"], true);
incomeTable.Name = "tblIncome";
incomeTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];

ITable expensesTable = worksheet.Tables.Add(worksheet.Range["B10:C23"], true);
expensesTable.Name = "tblExpenses";
expensesTable.TableStyle = workbook.TableStyles["TableStyleMedium4"];
```

## 设置计算公式

创建两个自定义名称来对月度的收入和支出分别做求和，然后在单元格中使用所定义的名称来设置公式，计算月度的收入总和以及支出总和，再分别设置另外两个公式，一个公式来计算支出占收入的百分比，一个来计算月度的净收入。

```csharp
worksheet.Names.Add("TotalMonthlyIncome", "=SUM(tblIncome[AMOUNT])");
worksheet.Names.Add("TotalMonthlyExpenses", "=SUM(tblExpenses[AMOUNT])");

worksheet.Range["E3"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G3"].Formula = "=TotalMonthlyExpenses/TotalMonthlyIncome";
worksheet.Range["G6"].Formula = "=TotalMonthlyIncome";
worksheet.Range["G7"].Formula = "=TotalMonthlyExpenses";
worksheet.Range["G9"].Formula = "=TotalMonthlyIncome-TotalMonthlyExpenses";
```


## 自定义样式

有两种方式来给单元格或区域设置样式：
- 给单元格或区域应用一个内置的或者自定义的命名样式
- 直接设置单元格或区域的某个具体的样式属性

在这里，我们会修改内置的一些命名样式--"Currency" "Heading 1"和"Percent"，然后把它们再应用到某个区域或单元格，而对其它的区域，我们则直接修改区域的样式属性。

```csharp
IStyle currencyStyle = workbook.Styles["Currency"];
currencyStyle.IncludeAlignment = true;
currencyStyle.HorizontalAlignment = HorizontalAlignment.Left;
currencyStyle.VerticalAlignment = VerticalAlignment.Bottom;
currencyStyle.NumberFormat = "$#,##0.00";

IStyle heading1Style = workbook.Styles["Heading 1"];
heading1Style.IncludeAlignment = true;
heading1Style.HorizontalAlignment = HorizontalAlignment.Center;
heading1Style.VerticalAlignment = VerticalAlignment.Center;
heading1Style.Font.Name = "Century Gothic";
heading1Style.Font.Bold = true;
heading1Style.Font.Size = 11;
heading1Style.Font.Color = Color.White;
heading1Style.IncludeBorder = false;
heading1Style.IncludePatterns = true;
heading1Style.Interior.Color = Color.FromRGB(32, 61, 64);

IStyle percentStyle = workbook.Styles["Percent"];
percentStyle.IncludeAlignment = true;
percentStyle.HorizontalAlignment = HorizontalAlignment.Center;
percentStyle.IncludeFont = true;
percentStyle.Font.Color = Color.FromRGB(32, 61, 64);
percentStyle.Font.Name = "Century Gothic";
percentStyle.Font.Bold = true;
percentStyle.Font.Size = 14;

worksheet.SheetView.DisplayGridlines = false;
worksheet.Range["C4:C7, C11:C23, G6:G7, G9"].Style = currencyStyle;
worksheet.Range["B2, B9, E2, E5"].Style = heading1Style;
worksheet.Range["G3"].Style = percentStyle;

worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E6:G6"].Borders[BordersIndex.EdgeBottom].Color = Color.FromRGB(32, 61, 64);
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].LineStyle = BorderLineStyle.Medium;
worksheet.Range["E7:G7"].Borders[BordersIndex.EdgeBottom].Color = Color.FromRGB(32, 61, 64);

worksheet.Range["E9:G9"].Interior.Color = Color.FromRGB(32, 61, 64);
worksheet.Range["E9:G9"].HorizontalAlignment = HorizontalAlignment.Left;
worksheet.Range["E9:G9"].VerticalAlignment = VerticalAlignment.Center;
worksheet.Range["E9:G9"].Font.Name = "Century Gothic";
worksheet.Range["E9:G9"].Font.Bold = true;
worksheet.Range["E9:G9"].Font.Size = 11;
worksheet.Range["E9:G9"].Font.Color = Color.White;
worksheet.Range["E3:F3"].Borders.Color = Color.FromRGB(32, 61, 64);
```


## 添加条件格式

Spread表格组件支持Excel所有类型的条件格式，在这里我们会创建一个渐变的数据条条件格式用以可视化地展示支出收入比，并将它设置为只显示数据条而不显示单元格的值。

```csharp
IDataBar dataBar = worksheet.Range["E3"].FormatConditions.AddDatabar();
dataBar.MinPoint.Type = ConditionValueTypes.Number;
dataBar.MinPoint.Value = 1;
dataBar.MaxPoint.Type = ConditionValueTypes.Number;
dataBar.MaxPoint.Value = "=TotalMonthlyIncome";
dataBar.BarFillType = DataBarFillType.Gradient;
dataBar.BarColor.Color = Color.Red;
dataBar.ShowValue = false;
```

## 添加图表

创建一个柱状图来展示收入和支出的差距，为了让图表更加美观，我们会设置数据系列，图表区域，轴线, 轴刻度以及数据点的样式。

```csharp
IShape shape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 339, 247, 316.5, 346);
shape.Chart.ChartArea.Format.Line.Transparency = 1;
shape.Chart.ColumnGroups[0].Overlap = 0;
shape.Chart.ColumnGroups[0].GapWidth = 37;

IAxis category_axis = shape.Chart.Axes.Item(AxisType.Category);
category_axis.Format.Line.Color.RGB = Color.Black;
category_axis.TickLabels.Font.Size = 11;
category_axis.TickLabels.Font.Color.RGB = Color.Black;

IAxis series_axis = shape.Chart.Axes.Item(AxisType.Value);
series_axis.Format.Line.Weight = 1;
series_axis.Format.Line.Color.RGB = Color.Black;
series_axis.TickLabels.NumberFormat = "$###0";
series_axis.TickLabels.Font.Size = 11;
series_axis.TickLabels.Font.Color.RGB = Color.Black;

ISeries chartSeries = shape.Chart.SeriesCollection.NewSeries();
chartSeries.Formula = "=SERIES(\"Simple Budget\",{\"Income\",\"Expenses\"},'Sheet1'!$G$6:$G$7,1)";
chartSeries.Points[0].Format.Fill.Color.RGB = Color.FromRGB(176, 21, 19);
chartSeries.Points[1].Format.Fill.Color.RGB = Color.FromRGB(234, 99, 18);
chartSeries.DataLabels.Font.Size = 11;
chartSeries.DataLabels.Font.Color.RGB = Color.Black;
chartSeries.DataLabels.ShowValue = true;
chartSeries.DataLabels.Position = DataLabelPosition.OutsideEnd;

```

## 保存为Excel

最后，我们会将Workbook保存为一个名为SimpleBudget.xlsx的Excel文件。

```csharp
workbook.Save(@"SimpleBudget.xlsx");
```

此时，你可以下载并查看保存的[SimpleBudget.xlsx](api/examples/xlsx/GrapeCity.Documents.Excel.Examples.Tutorial?fileName=SimpleBudget). 同时，你也可以下载本教程的[源代码工程](GrapeCity.Documents.Excel.Tutorial.zip)并亲自运行它，在你尝试运行源代码工程之前请先确保你已经安装了[.NET Core](https://www.microsoft.com/net/core)。