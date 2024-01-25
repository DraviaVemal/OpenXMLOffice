---
layout:
  title:
    visible: true
  description:
    visible: false
  tableOfContents:
    visible: true
  outline:
    visible: true
  pagination:
    visible: true
---

# Chart

The `Chart` class, a versatile component within the `OpenXMLOffice.Presentation` library, empowers developers to seamlessly integrate various types of charts into PowerPoint presentations. This class supports multiple chart types and configurations, allowing users to add new charts to a slide or replace existing shapes with dynamic and data-driven visualizations.

<details>

<summary>List of supported charts</summary>

* **Column Chart:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* **Line Chart:**
  * Cluster
  * Stacked
  * 100% Stacked
  * Cluster Marker
  * Stacked Marker
  * 100% Stacked Marker

<!---->

* **Pie Chart:**
  * Pie
  * Doughnut

<!---->

* **Bar Chart:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* **Area Chart:**
  * Cluster
  * Stacked
  * 100% Stacked

<!---->

* **X Y (Scatter) Chart:**
  * Scatter
  * Scatter Smooth Line Marker
  * Scatter Smooth Line
  * Scatter Line Marker
  * Scatter Line
  * Bubble

</details>

### Basic Code Samples

&#x20;For each chart family `ChartSetting` have its releavent options and settings for customization.

```csharp
public void ChartSample(PowerPoint powerPoint)
{
    // Default Chart Type
    powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting());
    // Customised Chart Type
    powerPoint.GetSlideByIndex(0).AddChart(CreateDataCellPayload(), new AreaChartSetting()
    {
        AreaChartTypes = AreaChartTypes.STACKED
    });
    Slide slide = powerPoint.GetSlideByIndex(1);
    Shape shape = slide.FindShapeByText("shape_id_1");
    shape.ReplaceChart(new Chart(slide, CreateDataCellPayload(),
            new BarChartSetting()
            {
                ChartLegendOptions = new ChartLegendOptions()
                {
                    LegendPosition = ChartLegendOptions.LegendPositionValues.RIGHT
                }
            })
}
```

### `ChartSetting` Options

<table><thead><tr><th width="250">Property</th><th width="212">Type</th><th>Details</th></tr></thead><tbody><tr><td>chartDataSetting</td><td><a href="./#chartdatasetting-options">ChartDataSetting</a></td><td>This setting enables users to customize both the input chart data range and value from cell labels with precision.</td></tr><tr><td>chartGridLinesOptions</td><td><a href="./#chartgridlinesoptions-options">ChartGridLinesOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>chartLegendOptions</td><td><a href="./#chartlegendoptions-options">ChartLegendOptions</a></td><td>This feature offers crisp options for users to finely customize the gridline settings of the chart.</td></tr><tr><td>height</td><td>uint</td><td>This parameter precisely determines the height of the entire chart. Default : 6858000</td></tr><tr><td>width</td><td>uint</td><td>This parameter precisely determines the width of the entire chart. Default : 12192000</td></tr><tr><td>x</td><td>uint</td><td>This parameter precisely determines the X position of the entire chart. Default: 0</td></tr><tr><td>y</td><td>uint</td><td>This parameter precisely determines the Y position of the entire chart. Default : 0</td></tr></tbody></table>

### `ChartDataSetting` Options

<table><thead><tr><th width="293">Property</th><th width="69">Type</th><th>Details</th></tr></thead><tbody><tr><td>chartDataColumnEnd</td><td>uint</td><td>Specify the number of columns for chart series; set to 0 for utilizing all columns. <br>Default: 0</td></tr><tr><td>chartDataColumnStart</td><td>uint</td><td>Specify the starting column for chart data.<br>Default: 0</td></tr><tr><td>chartDataRowEnd</td><td>uint</td><td>Specify the number of rows for chart series; set to 0 for utilizing all rows. <br>Default: 0</td></tr><tr><td>chartDataRowStart</td><td>uint</td><td>Specify the starting row for chart data.<br>Default: 0</td></tr></tbody></table>

### `ChartGridLinesOptions` Options

<table><thead><tr><th width="276">Property</th><th width="91">Type</th><th>Details</th></tr></thead><tbody><tr><td>isMajorCategoryLinesEnabled</td><td>bool</td><td>Toggle visibility of major category lines with clarity.</td></tr><tr><td>isMajorValueLinesEnabled</td><td>bool</td><td>Toggle visibility of major value lines with clarity.</td></tr><tr><td>isMinorCategoryLinesEnabled</td><td>bool</td><td>Toggle visibility of minor category lines with clarity.</td></tr><tr><td>isMinorValueLinesEnabled</td><td>bool</td><td>Toggle visibility of minor value lines with clarity.</td></tr></tbody></table>

### `ChartLegendOptions` Options

| Property             | Type                 | Details                                                                              |
| -------------------- | -------------------- | ------------------------------------------------------------------------------------ |
| isEnableLegend       | bool                 | Toggle visibility of legend with clarity.                                            |
| isLegendChartOverLap | bool                 | Activate the option for a sleek and tidy display by allowing the legends to overlap. |
| isBold               | bool                 | Provide the option to set text in a bold format with clarity.                        |
| isItalic             | bool                 | Provide the option to set text in a italic format with clarity.                      |
| fontSize             | float                | Provide the option to set font size with clarity.                                    |
| fontColor            | string?              | Optional font color using hex code (without #), defaulting to Theme Text 1.          |
| underLineValues      | UnderLineValues      | Text underline options. Default: None                                                |
| strikeValues         | StrikeValues         | Text strike options                                                                  |
| legendPosition       | LegendPositionValues | Legend position in chart. Default: Bottom                                            |
