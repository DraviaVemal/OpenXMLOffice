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

```csharp
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.AreaChartSetting()
	{
		areaChartTypes = G.AreaChartTypes.STACKED,
		chartAxesOptions = new()
		{
			horizontalFontSize = 20,
			verticalFontSize = 25
		}
	});
```

### `AreaChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to area chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>areaChartDataLabel</td><td><a href="area.md#areachartdatalabel-options">AreaChartDataLabel</a></td><td></td></tr><tr><td>areaChartSeriesSettings</td><td>List&#x3C;<a href="area.md#areachartseriessetting-options">AreaChartSeriesSetting</a>?></td><td></td></tr><tr><td>areaChartTypes</td><td>AreaChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td></td></tr></tbody></table>

### `AreaChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to area chart.

| Property          | Type                    | Details |
| ----------------- | ----------------------- | ------- |
| dataLabelPosition | DataLabelPositionValues |         |

### `AreaChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to area chart.

|                    |                                                          |   |
| ------------------ | -------------------------------------------------------- | - |
| areaChartDataLabel | [AreaChartDataLabel](area.md#areachartdatalabel-options) |   |
| fillColor          | string?                                                  |   |
