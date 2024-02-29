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

# Scatter

Add chart method present in slide component or you can replace the chart using shape componenet.

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ScatterChartSetting());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ScatterChartSetting()
	{
		scatterChartTypes = G.ScatterChartTypes.BUBBLE
	});
```
{% endtab %}
{% endtabs %}

### `ScatterChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to scatter chart.

<table><thead><tr><th width="250">Property</th><th width="287">Type</th><th>Details</th></tr></thead><tbody><tr><td>scatterChartDataLabel</td><td><a href="scatter.md#scatterchartdatalabel-options">ScatterChartDataLabel</a></td><td></td></tr><tr><td>scatterChartSeriesSettings</td><td>List&#x3C;<a href="scatter.md#scatterchartseriessetting-options">ScatterChartSeriesSetting</a>?></td><td></td></tr><tr><td>scatterChartTypes</td><td>ScatterChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td>ChartAxesOptions</td><td></td></tr></tbody></table>

### `ScatterChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to scatter chart.

<table><thead><tr><th width="194">Property</th><th width="220">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td></td></tr></tbody></table>

### `ScatterChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to scatter chart.

<table><thead><tr><th width="226">Property</th><th width="209">Type</th><th>Details</th></tr></thead><tbody><tr><td>ScatterChartDataLabel</td><td><a href="scatter.md#scatterchartdatalabel-options">ScatterChartDataLabel</a></td><td></td></tr></tbody></table>
