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

# Bar

Add chart method present in slide component or you can replace the chart using shape componenet.

### Basic Code Sample

```csharp
// Bare minimum
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.BarChartSetting());
// Some additional samples
powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK)
	.AddChart(CreateDataCellPayload(), new G.BarChartSetting()
	{
		chartAxesOptions = new()
		{
			isHorizontalAxesEnabled = false,
		},
		barChartDataLabel = new G.BarChartDataLabel()
		{
			dataLabelPosition = G.BarChartDataLabel.DataLabelPositionValues.INSIDE_END,
			showValue = true,
		},
		barChartSeriesSettings = new(){
			new(),
			new(){
				barChartDataLabel = new G.BarChartDataLabel(){
					dataLabelPosition = G.BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END,
					showCategoryName= true
				}
			}
		}
	});
```

### `BarChartSetting` Options

Contains options details extended from [`ChartSetting`](./#chartsetting-options) that are specific to bar chart.

<table><thead><tr><th width="238">Property</th><th width="262">Type</th><th>Details</th></tr></thead><tbody><tr><td>barChartDataLabel</td><td><a href="bar.md#barchartdatalabel-options">BarChartDataLabel</a></td><td></td></tr><tr><td>barChartSeriesSettings</td><td>List&#x3C;<a href="bar.md#barchartseriessetting-options">BarChartSeriesSetting</a>?></td><td></td></tr><tr><td>barChartTypes</td><td>BarChartTypes</td><td></td></tr><tr><td>chartAxesOptions</td><td><a href="./#chartaxesoptions-options">ChartAxesOptions</a></td><td></td></tr><tr><td>barGraphicsSetting</td><td><a href="bar.md#bargraphicssetting-options">BarGraphicsSetting</a></td><td></td></tr></tbody></table>

### `BarChartDataLabel` Options

Contains options details extended from [`ChartDataLabel`](./#chartdatalabel-options) that are specific to bar chart.

| Property          | Type                    | Details |
| ----------------- | ----------------------- | ------- |
| dataLabelPosition | DataLabelPositionValues |         |

### `barChartSeriesSetting` Options

Contains options details extended from [`ChartSeriesSetting`](./#chartseriessetting-options) that are specific to bar chart.

|                   |                                                       |   |
| ----------------- | ----------------------------------------------------- | - |
| barChartDataLabel | [BarChartDataLabel](bar.md#barchartdatalabel-options) |   |
| fillColor         | string?                                               |   |

### `BarGraphicsSetting` Options

| Property    | Type | Details |
| ----------- | ---- | ------- |
| categoryGap | int  |         |
| seriesGap   | int  |         |
