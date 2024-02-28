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

# Column

Add chart method present in slide component or you can replace the chart using shape componenet.

### Basic Code Sample

```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ColumnChartSetting());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.ColumnChartSetting()
	{
		titleOptions = new()
		{
			title = "Column Chart"
		},
		chartLegendOptions = new G.ChartLegendOptions()
		{
			legendPosition = G.ChartLegendOptions.LegendPositionValues.TOP,
			fontSize = 5
		},
		columnChartSeriesSettings = new(){
			null,
			new(){
				columnChartDataPointSettings = new(){
				null,
				new(){
					fillColor = "FF0000"
				},
				new(){
					fillColor = "00FF00"
				},
			},
				fillColor= "AABBCC"
			},
			new(){
				fillColor= "CCBBAA"
			}
		}
	});
```

### `ColumnChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to column chart.

<table><thead><tr><th width="251">Property</th><th width="287">Type</th><th>Details</th></tr></thead><tbody><tr><td>columnChartDataLabel</td><td><a href="column.md#columnchartdatalabel-options">ColumnChartDataLabel</a></td><td></td></tr><tr><td>columnChartSeriesSettings</td><td>List&#x3C;<a href="column.md#columnchartsetting-options">ColumnChartSeriesSetting</a>?></td><td></td></tr><tr><td>columnChartTypes</td><td>ColumnChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td></td></tr><tr><td>columnGraphicsSetting</td><td><a href="column.md#columngraphicssetting-options">ColumnGraphicsSetting</a></td><td></td></tr></tbody></table>

### `ColumnChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to column chart.

| Property          | Type                    | Details |
| ----------------- | ----------------------- | ------- |
| dataLabelPosition | DataLabelPositionValues |         |

### `ColumnChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to column chart.

|                      |                                                                |   |
| -------------------- | -------------------------------------------------------------- | - |
| columnChartDataLabel | [ColumnChartDataLabel](column.md#columnchartdatalabel-options) |   |
| fillColor            | string?                                                        |   |

### `ColumnGraphicsSetting` Options

| Property    | Type | Details |
| ----------- | ---- | ------- |
| categoryGap | int  |         |
| seriesGap   | int  |         |
