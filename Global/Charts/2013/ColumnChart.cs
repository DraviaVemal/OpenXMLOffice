// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the settings for a column chart.
	/// </summary>
	public class ColumnChart : ChartBase
	{
		private const int DefaultGapWidth = 150;
		private const int DefaultOverlap = 100;

		/// <summary>
		/// Column Chart Setting
		/// </summary>
		protected ColumnChartSetting columnChartSetting;

		internal ColumnChart(ColumnChartSetting columnChartSetting) : base(columnChartSetting)
		{
			this.columnChartSetting = columnChartSetting;
		}

		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		public ColumnChart(ColumnChartSetting columnChartSetting, ChartData[][] dataCols) : base(columnChartSetting)
		{
			this.columnChartSetting = columnChartSetting;
			SetChartPlotArea(CreateChartPlotArea(dataCols));
		}

		private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
		{
			C.PlotArea plotArea = new();
			plotArea.Append(CreateLayout(columnChartSetting.plotAreaOptions?.manualLayout));
			plotArea.Append(CreateColumnChart(CreateDataSeries(dataCols, columnChartSetting.chartDataSetting)));
			plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
			{
				id = CategoryAxisId,
				crossAxisId = ValueAxisId,
				fontSize = columnChartSetting.chartAxesOptions.horizontalFontSize,
				isBold = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = columnChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
				invertOrder = columnChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
			}));
			plotArea.Append(CreateValueAxis(new ValueAxisSetting()
			{
				id = ValueAxisId,
				crossAxisId = CategoryAxisId,
				fontSize = columnChartSetting.chartAxesOptions.verticalFontSize,
				isBold = columnChartSetting.chartAxesOptions.isVerticalBold,
				isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
				isVisible = columnChartSetting.chartAxesOptions.isVerticalAxesEnabled,
				invertOrder = columnChartSetting.chartAxesOptions.invertVerticalAxesOrder,
			}));
			plotArea.Append(CreateChartShapeProperties());
			return plotArea;
		}

		internal C.BarChart CreateColumnChart(List<ChartDataGrouping> chartDataGroupings)
		{
			C.BarChart columnChart = new(
				new C.BarDirection { Val = C.BarDirectionValues.Column },
				new C.BarGrouping
				{
					Val = columnChartSetting.columnChartTypes switch
					{
						ColumnChartTypes.STACKED => C.BarGroupingValues.Stacked,
						ColumnChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
						// Clusted
						_ => C.BarGroupingValues.Clustered,
					}
				},
				new C.VaryColors { Val = false });
			int seriesIndex = 0;
			chartDataGroupings.ForEach(Series =>
			{
				columnChart.Append(CreateColumnChartSeries(seriesIndex, Series));
				seriesIndex++;
			});
			if (columnChartSetting.columnChartTypes == ColumnChartTypes.CLUSTERED)
			{
				columnChart.Append(new C.GapWidth { Val = (UInt16Value)columnChartSetting.columnGraphicsSetting.categoryGap });
				columnChart.Append(new C.Overlap { Val = (SByteValue)columnChartSetting.columnGraphicsSetting.seriesGap });
			}
			else
			{
				columnChart.Append(new C.GapWidth { Val = DefaultGapWidth });
				columnChart.Append(new C.Overlap { Val = DefaultOverlap });
			}
			C.DataLabels? dataLabels = CreateColumnDataLabels(columnChartSetting.columnChartDataLabel);
			if (dataLabels != null)
			{
				columnChart.Append(dataLabels);
			}
			columnChart.Append(new C.AxisId { Val = CategoryAxisId });
			columnChart.Append(new C.AxisId { Val = ValueAxisId });
			return columnChart;
		}

		private C.BarChartSeries CreateColumnChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
		{
			SolidFillModel GetSeriesFillColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = columnChartSetting.columnChartSeriesSettings?
							.Select(item => item?.fillColor)
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
			SolidFillModel GetSeriesBorderColor()
			{
				SolidFillModel solidFillModel = new();
				string? hexColor = columnChartSetting.columnChartSeriesSettings?
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
			ShapePropertiesModel shapePropertiesModel = new()
			{
				solidFill = GetSeriesFillColor(),
				outline = new()
				{
					solidFill = GetSeriesBorderColor()
				}
			};
			C.DataLabels? dataLabels = seriesIndex < columnChartSetting.columnChartSeriesSettings.Count ?
				CreateColumnDataLabels(columnChartSetting.columnChartSeriesSettings[seriesIndex]?.columnChartDataLabel ?? new ColumnChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
			C.BarChartSeries series = new(
				new C.Index { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.Order { Val = new UInt32Value((uint)chartDataGrouping.id) },
				new C.InvertIfNegative { Val = true },
				CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
			series.Append(CreateChartShapeProperties(shapePropertiesModel));
			int dataPointCount = columnChartSetting.columnChartSeriesSettings?.ElementAtOrDefault(seriesIndex)?.columnChartDataPointSettings.Count ?? 0;
			for (uint index = 0; index < dataPointCount; index++)
			{
				if (columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings != null &&
				index < columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings.Count &&
				columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings[(int)index] != null)
				{
					SolidFillModel GetDataPointFill()
					{
						SolidFillModel solidFillModel = new();
						string? hexColor = columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings?
									.Select(item => item?.fillColor)
									.ToList().ElementAtOrDefault((int)index);
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
					SolidFillModel GetDataPointBorder()
					{
						SolidFillModel solidFillModel = new();
						string? hexColor = columnChartSetting.columnChartSeriesSettings?[seriesIndex]?.columnChartDataPointSettings?
									.Select(item => item?.borderColor)
									.ToList().ElementAtOrDefault((int)index);
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
					C.DataPoint dataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
					dataPoint.Append(CreateChartShapeProperties(new ShapePropertiesModel()
					{
						solidFill = GetDataPointFill(),
						outline = new()
						{
							solidFill = GetDataPointBorder()
						}
					}));
					series.Append(dataPoint);
				}
			}
			if (dataLabels != null)
			{
				series.Append(dataLabels);
			}
			series.Append(CreateCategoryAxisData(chartDataGrouping.xAxisFormula!, chartDataGrouping.xAxisCells!));
			series.Append(CreateValueAxisData(chartDataGrouping.yAxisFormula!, chartDataGrouping.yAxisCells!));
			if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
			{
				series.Append(new C.ExtensionList(new C.Extension(
					CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
				)
				{ Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
			}
			return series;
		}

		private C.DataLabels? CreateColumnDataLabels(ColumnChartDataLabel columnChartDataLabel, int? dataLabelCounter = 0)
		{
			if (columnChartDataLabel.showValue || columnChartDataLabel.showValueFromColumn || columnChartDataLabel.showCategoryName || columnChartDataLabel.showLegendKey || columnChartDataLabel.showSeriesName)
			{
				C.DataLabels dataLabels = CreateDataLabels(columnChartDataLabel, dataLabelCounter);
				dataLabels.InsertAt(new C.DataLabelPosition()
				{
					Val = columnChartDataLabel.dataLabelPosition switch
					{
						ColumnChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
						ColumnChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
						ColumnChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
						_ => C.DataLabelPositionValues.Center
					}
				}, 0);
				return dataLabels;
			}
			return null;
		}


	}
}
