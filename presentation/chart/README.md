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

### Code Samples

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

