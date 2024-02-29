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

{% tabs %}
{% tab title="C#" %}
<pre class="language-csharp"><code class="lang-csharp"><strong>// Bare minimum
</strong>owerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
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
{% endtab %}
{% endtabs %}

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

<table><thead><tr><th width="258">Property</th><th width="292">Type</th><th>Details</th></tr></thead><tbody><tr><td>lineChartDataLabel</td><td><a href="line.md#linechartdatalabel-options">LineChartDataLabel</a></td><td></td></tr><tr><td>lineChartLineFormat</td><td><a href="line.md#linechartlineformat-options">LineChartLineFormat</a></td><td></td></tr><tr><td>lineChartDataPointSettings</td><td>List&#x3C;LineChartDataPointSetting?></td><td>TODO</td></tr></tbody></table>

### `LineChartLineFormat` Options

<table><thead><tr><th width="213">Property</th><th width="270">Type</th><th>Details</th></tr></thead><tbody><tr><td>transparency</td><td>int?</td><td></td></tr><tr><td>width</td><td>int?</td><td></td></tr><tr><td>outlineCapTypeValues</td><td>OutlineCapTypeValues?</td><td></td></tr><tr><td>outlineLineTypeValues</td><td>OutlineLineTypeValues?</td><td></td></tr><tr><td>beginArrowValues</td><td>DrawingBeginArrowValues?</td><td></td></tr><tr><td>endArrowValues</td><td>DrawingEndArrowValues?</td><td></td></tr><tr><td>dashType</td><td>DrawingPresetLineDashValues?</td><td></td></tr><tr><td>lineStartWidth</td><td>LineWidthValues?</td><td></td></tr><tr><td>lineEndWidth</td><td>LineWidthValues?</td><td></td></tr></tbody></table>
