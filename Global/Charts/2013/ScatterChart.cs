// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the types of scatter charts.
	/// </summary>
	public class ScatterChart : ChartBase
	{


		/// <summary>
		/// Scatter Chart Setting
		/// </summary>
		protected ScatterChartSetting scatterChartSetting;

		internal ScatterChart(ScatterChartSetting scatterChartSetting) : base(scatterChartSetting)
		{
			this.scatterChartSetting = scatterChartSetting;
		}

		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		public ScatterChart(ScatterChartSetting scatterChartSetting, ChartData[][] dataCols) : base(scatterChartSetting)
		{
			this.scatterChartSetting = scatterChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
		{
			if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
			{
				scatterChartSetting.chartDataSetting.is3Ddata = true;
				if ((dataCols.Length - 1) % 2 != 0)
				{
					throw new ArgumentOutOfRangeException("Required 3D Data Size is not met.");
				}
			}
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(scatterChartSetting.plotAreaOptions?.manualLayout));
			plotArea.Append(scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE ?
				CreateChart<C.BubbleChart>(CreateDataSeries(dataCols, scatterChartSetting.chartDataSetting)) :
				CreateChart<C.ScatterChart>(CreateDataSeries(dataCols, scatterChartSetting.chartDataSetting)));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = CategoryAxisId,
				axisPosition = AxisPosition.BOTTOM,
				crossAxisId = ValueAxisId,
				fontSize = scatterChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = scatterChartSetting.chartAxesOptions.isHorizontalBold,
				isItalic = scatterChartSetting.chartAxesOptions.isHorizontalItalic,
				isVisible = scatterChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = scatterChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = scatterChartSetting.chartAxesOptions.verticalFontSize,
				isBold = scatterChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = scatterChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = scatterChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = scatterChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal T CreateChart<T>(List<ChartDataGrouping> chartDataGroupings) where T : new()
		{
			T chartType = new();
			if (chartType is C.ScatterChart scatterChart)
			{
				scatterChart.Append(new C.ScatterStyle
				{
					Val = scatterChartSetting.scatterChartTypes switch
					{
						ScatterChartTypes.SCATTER_SMOOTH => C.ScatterStyleValues.Smooth,
						ScatterChartTypes.SCATTER_SMOOTH_MARKER => C.ScatterStyleValues.SmoothMarker,
						ScatterChartTypes.SCATTER_STRIGHT => C.ScatterStyleValues.Line,
						ScatterChartTypes.SCATTER_STRIGHT_MARKER => C.ScatterStyleValues.LineMarker,
						// Clusted
						_ => C.ScatterStyleValues.LineMarker,
					}
				});
			}
			if (chartType is OpenXmlCompositeElement chart)
			{
				chart.Append(new C.VaryColors() { Val = false });
				int seriesIndex = 0;
				chartDataGroupings.ForEach(Series =>
				{
					chart.Append(CreateScatterChartSeries(seriesIndex, Series));
					seriesIndex++;
				});
				C.DataLabels? dataLabels = CreateScatterDataLabels(scatterChartSetting.scatterChartDataLabel);
				if (dataLabels != null)
				{
					chart.Append(dataLabels);
				}
				if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
				{
					chart.Append(new C.BubbleScale() { Val = 100 });
					chart.Append(new C.ShowNegativeBubbles() { Val = true });
				}
				chart.Append(new C.AxisId { Val = CategoryAxisId });
				chart.Append(new C.AxisId { Val = ValueAxisId });
			}
			return chartType;
		}

		private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			C.DataLabels? dataLabels = seriesIndex < scatterChartSetting.scatterChartSeriesSettings.Count ?
				CreateScatterDataLabels(scatterChartSetting.scatterChartSeriesSettings?[seriesIndex]?.scatterChartDataLabel ?? new ScatterChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			SolidFillModel GetSeriesBorderColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = scatterChartSetting.scatterChartSeriesSettings?
							.Select(item => item?.borderColor)
							.ToList().ElementAtOrDefault(seriesIndex);
				if (hexColor != null)
				{
					solidFillModel.hexColor = hexColor;
					return solidFillModel;
				}
				else
				{
					solidFillModel.schemeColorModel = new()
					{
						themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
					};
				}
				return solidFillModel;
			}
			MarkerModel markerModel = new();
			if (new[] { ScatterChartTypes.SCATTER, ScatterChartTypes.SCATTER_SMOOTH_MARKER, ScatterChartTypes.SCATTER_STRIGHT_MARKER }.Contains(scatterChartSetting.scatterChartTypes))
			{
				markerModel.markerShapeValues = scatterChartSetting.scatterChartTypes == ScatterChartTypes.SCATTER ? MarkerModel.MarkerShapeValues.AUTO : MarkerModel.MarkerShapeValues.CIRCLE;
				markerModel.shapeProperties = new()
				{
					solidFill = new()
					{
						schemeColorModel = new()
						{
							themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
						}
					},
					outline = new()
					{
						solidFill = new()
						{
							schemeColorModel = new()
							{
								themeColorValues = ThemeColorValues.ACCENT_1 + (chartDataGrouping.id % AccentColurCount),
							}
						}
					}
				};
			}
			C.ScatterChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			ShapePropertiesModel shapePropertiesModel = new()
			{
				outline = new()
				{
					solidFill = scatterChartSetting.scatterChartTypes == ScatterChartTypes.SCATTER ? null : GetSeriesBorderColor(),
				}
			};
			if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
			{
				shapePropertiesModel.solidFill = new()
				{
					schemeColorModel = new()
					{
						themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
						tint = 75000,

					}
				};
				series.Append(new C.InvertIfNegative() { Val = false });
			}
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			if (scatterChartSetting.scatterChartTypes != ScatterChartTypes.BUBBLE)
			{
				series.Append(CreateMarker(markerModel));
			}
			if (dataLabels != null)
			{
				series.Append(dataLabels);
			}
			series.Append(CreateXValueAxisData(chartDataGrouping.xAxisFormula!, chartDataGrouping.xAxisCells!));
			series.Append(CreateYValueAxisData(chartDataGrouping.yAxisFormula!, chartDataGrouping.yAxisCells!));
			if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
			{
				series.Append(CreateBubbleSizeAxisData(chartDataGrouping.zAxisFormula!, chartDataGrouping.zAxisCells!));
				series.Append(new C.Bubble3D() { Val = false });
			}
			else
			{
				series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(scatterChartSetting.scatterChartTypes) });
			}
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}

		private C.DataLabels? CreateScatterDataLabels(ScatterChartDataLabel scatterChartDataLabel, int? dataLabelCounter = 0)
		{
			if (scatterChartDataLabel.showValue || scatterChartDataLabel.showValueFromColumn || scatterChartDataLabel.showCategoryName || scatterChartDataLabel.showLegendKey || scatterChartDataLabel.showSeriesName || scatterChartDataLabel.showBubbleSize)
			{
				C.DataLabels dataLabels = CreateDataLabels(scatterChartDataLabel, dataLabelCounter);
				dataLabels.Append(new C.ShowBubbleSize { Val = scatterChartDataLabel.showBubbleSize });
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = scatterChartDataLabel.dataLabelPosition switch
					{
						ScatterChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
						ScatterChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
						ScatterChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
						ScatterChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
						//Center
						_ => C.DataLabelPositionValues.Center,
					}
				}, 0);
				return dataLabels;
			}
			return null;
		}


	}
}
