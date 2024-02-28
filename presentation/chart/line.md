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

# Line

Add chart method present in slide component or you can replace the chart using shape componenet.

### Basic Code Sample

<pre class="language-csharp"><code class="lang-csharp">// Bare minimum
owerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.LineChartSetting());
// Some additional samples
<strong>powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
</strong>	.AddChart(CreateDataCellPayload(), new G.LineChartSetting()
	{
		lineChartSeriesSettings = new(){
			new(){
				lineChartLineFormat = new(){
					dashType = G.DrawingPresetLineDashValues.DASH_DOT,
					lineColor = "FF0000",
					beginArrowValues= G.DrawingBeginArrowValues.ARROW,
					endArrowValues= G.DrawingEndArrowValues.TRIANGLE,
					lineStartWidth = G.LineWidthValues.MEDIUM,
					lineEndWidth = G.LineWidthValues.LARGE,
					outlineCapTypeValues = G.OutlineCapTypeValues.ROUND,
					outlineLineTypeValues = G.OutlineLineTypeValues.DOUBLE,
					width = 5
				}
			}
		}
	});
</code></pre>

### `LineChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to line chart.

<table><thead><tr><th width="251">Property</th><th width="287">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td></td></tr><tr><td>lineChartSeriesSettings</td><td>List&#x3C;<a href="line.md#linechartseriessetting-options">LineChartSeriesSetting</a>?></td><td></td></tr><tr><td>lineChartTypes</td><td>LineChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td></td></tr></tbody></table>

### `LineChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to line chart.

| Property          | Type                    | Details |
| ----------------- | ----------------------- | ------- |
| dataLabelPosition | DataLabelPositionValues |         |

### `LineChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to column chart.

|                    |                                                          |   |
| ------------------ | -------------------------------------------------------- | - |
| lineChartDataLabel | [LineChartDataLabel](line.md#linechartdatalabel-options) |   |
| fillColor          | string?                                                  |   |
