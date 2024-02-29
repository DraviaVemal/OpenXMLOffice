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

# Area

Add chart method present in slide component or you can replace the chart using shape componenet.

### Basic Code Sample

{% tabs %}
{% tab title="C#" %}
<pre class="language-csharp"><code class="lang-csharp">// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.AreaChartSetting());
<strong>// Some additional samples
</strong><strong>powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
</strong>	.AddChart(CreateDataCellPayload(), new G.AreaChartSetting()
	{
		areaChartTypes = G.AreaChartTypes.STACKED,
		chartAxesOptions = new()
		{
			horizontalFontSize = 20,
			verticalFontSize = 25
		}
	});
</code></pre>
{% endtab %}
{% endtabs %}

### `AreaChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to area chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>areaChartDataLabel</td><td><a href="area.md#areachartdatalabel-options">AreaChartDataLabel</a></td><td></td></tr><tr><td>areaChartSeriesSettings</td><td>List&#x3C;<a href="area.md#areachartseriessetting-options">AreaChartSeriesSetting</a>?></td><td></td></tr><tr><td>areaChartTypes</td><td>AreaChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td></td></tr></tbody></table>

### `AreaChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to area chart.

<table><thead><tr><th width="194">Property</th><th width="220">Type</th><th>Details</th></tr></thead><tbody><tr><td>dataLabelPosition</td><td>DataLabelPositionValues</td><td></td></tr></tbody></table>

### `AreaChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to area chart.

<table><thead><tr><th width="206"></th><th width="188"></th><th></th></tr></thead><tbody><tr><td>areaChartDataLabel</td><td><a href="area.md#areachartdatalabel-options">AreaChartDataLabel</a></td><td></td></tr><tr><td>fillColor</td><td>string?</td><td></td></tr></tbody></table>
